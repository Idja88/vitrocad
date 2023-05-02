param(
    [string]$url,
    [string]$importpath= ""
)

Connect-PnPOnline $url -CurrentCredentials

$text = Get-Content -Path $importpath
$rows = ConvertFrom-Csv -Delimiter "`t" -InputObject $text

#Определение всех уникальных подразделений и должностей
$subdivs = $rows | Select-Object Subdivision | Sort-Object Subdivision -Unique
$staffs = $rows | Select-Object Staffname | Sort-Object Staffname -Unique

#Загрузка подразделений и добавление их как группы Sharepoint
foreach ($subdiv in $subdivs){
    $SubValues =@{}
    $SubValues +=@{'Title' = $subdiv.Subdivision}

    #Проверка на существующие подразделения
    if(Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $subdiv.Subdivision}){
    Write-Host "Подразделение - " $subdiv.Subdivision " уже загружено." -ForegroundColor Green}
    else {
    Add-PnPListItem -List "Организационно-штатная структура" -ContentType "Подразделение" -Values $SubValues
    }
    
    #Проверка на существующие группы, как подразделения
    if($null -ne (Get-PnPGroup -Identity $subdiv.Subdivision)){
        Write-Host "Подразделение - " $subdiv.Subdivision "уже добавлено как группа." -ForegroundColor Green
    }
    else {
        New-PnPGroup -Title $subdiv.Subdivision -Owner $admin.UserName
    }
    
    $OssSubId = (Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $subdiv.Subdivision}).FieldValues.ID
    $SubGrp = (Get-PnPGroup -Identity $subdiv.Subdivision).Id
    
    $subdiv | Add-Member -MemberType NoteProperty -Name OssSubdivisionID -Value $OssSubId -Force
    $subdiv | Add-Member -MemberType NoteProperty -Name GroupID -Value $SubGrp -Force
}

#Загрузка должностей
foreach ($staff in $staffs){
    $StaffValues = @{}
    $StaffValues +=@{'Title' = $staff.Staffname}
    
    #Проверка на существующие должности
    if(Get-PnPListItem -List "Наименование должностей" | Where-Object {$_["Title"] -eq $staff.Staffname}){
    Write-Host "Должность - " $staff.Staffname " уже загружена." -ForegroundColor Green}
    else {
    Add-PnPListItem -List "Наименование должностей" -Values $StaffValues
    }
}

#Загрузка Физических Лиц, определение ID загруженных элементов и их занесение обратно в массив построчно, т.к для лукап поля нужны ID элементов.
foreach($row in $rows){
    $RowValues =@{}
    $RowValues +=@{'Title' = $Row.Surname}
    $RowValues +=@{'FirstName' = $Row.FirstName}
    $RowValues +=@{'MiddleName' = $Row.SecondName}
    $RowValues +=@{'FullName' = $Row.Fullname}
    $RowValues +=@{'Email' = $Row.Email}
    $RowValues +=@{'VitroOrgLogin' = $Row.Login}

    #Проверка на существующих пользователей, нужно проверять полное имя из-за однофамильцев
    if((Get-PnPListItem -List "Физические лица").FieldValues | Where-Object {$_.FullName -eq $Row.Fullname}){
    Write-Host "Пользователь - " $Row.Fullname " уже загружен." -ForegroundColor Green}
    else {
    Add-PnPlistItem -List "Физические лица" -Values $RowValues
    }
    
    $fiz = ((Get-PnPListItem -List "Физические лица").FieldValues | Where-Object {$_.FullName -eq $Row.Fullname}).ID
    $div = (Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $Row.Subdivision}).FieldValues.ID
    $stf = (Get-PnPListItem -List "Наименование должностей" | Where-Object {$_["Title"] -eq $Row.Staffname}).FieldValues.ID
    
    $row | Add-Member -MemberType NoteProperty -Name FizID -Value $fiz -Force
    $row | Add-Member -MemberType NoteProperty -Name SubdivID -Value $div -Force
    $row | Add-Member -MemberType NoteProperty -Name StaffID -Value $stf -Force
}

#Загрузка Сотрудников в ОШС по имеющимся ID 
foreach ($row in $rows){
    $RowValues =@{}
    $RowValues +=@{'Title' = $row.FullName}
    $RowValues +=@{'VitroOrgPerson' = $row.FizID}
    $RowValues +=@{'VitroOrgParentId'= $row.SubdivID}
    $RowValues +=@{'VitroOrgPositionName' = $row.StaffID}
    $RowValues +=@{'VitroOrgPositionDisplayName' = $row.Staffname+" - "+$row.FullName}
 
    #Проверка на существующих сотрудников
    if(Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $row.Fullname}){
    Write-Host "Cотрудник - " $row.Fullname " уже загружен." -ForegroundColor Green}
    else {
    Add-PnPListItem -List "Организационно-штатная структура" -ContentType "Сотрудник" -Values $RowValues
    }

    $oss = (Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $row.FullName}).FieldValues.ID

    $row | Add-Member -MemberType NoteProperty -Name OssID -Value $oss -Force
}

#Добавление пользователей в группы Sharepoint
foreach ($row in $rows){

    #Проверка на добавленных пользователей в группы
    if(Get-PnPGroupMembers -Identity $row.Subdivision | Where-Object {$_.LoginName -like ("*"+$row.Login)}){
    Write-Host "Пользователь - " $row.FullName "уже находится в группе"$row.Subdivision"." -ForegroundColor Green
    }
    else {
    Add-PnPUserToGroup -Identity $row.Subdivision -LoginName $row.Login
    Write-Host "Пользователь - " $row.FullName "добавлен в группу"$row.Subdivision"."
    }

    $grp = (Get-PnPGroup -Identity $row.Subdivision).Id

    $row | Add-Member -MemberType NoteProperty -Name GroupID -Value $grp -Force
 }

#Обновление подразделений ОШС - обновление поля Руководитель, т.к теперь загружены Сотрудники
$ruks = $rows | Select-Object Subdivision, FullName, Staffname, OssID | Where-Object Staffname -EQ "Начальник отдела"

foreach ($ruk in $ruks){
    $pod = (Get-PnPListItem -List "Организационно-штатная структура" | Where-Object {$_["Title"] -eq $ruk.Subdivision}).FieldValues.ID
    $RukValues =@{}
    $RukValues +=@{'VitroOrgStructureHeader' = $ruk.OssID}
  
    #Проверка на уже проставленных начальников отдела
    if((Get-PnPListItem -List "Организационно-штатная структура" -Id $pod).FieldValues.VitroOrgStructureHeader -eq $ruk.Fullname){
    Write-Host "У подразделения" $ruk.Subdivision "уже проставлен Начальник Отдела." -ForegroundColor Green
    }
    else{
    Set-PnPListItem -List "Организационно-штатная структура" -ContentType "Подразделение" -Identity $pod -Values $RukValues
    }
 }

#Обновление подразделений ОШС - обновление поля Группа
foreach ($subdiv in $subdivs){
    $GroupValues =@{}
    $GroupValues +=@{"VitroOrgLogin" = (""+$subdiv.GroupID)}

    #Проверка на уже проставленные группы
    if((Get-PnPListItem -List "Организационно-штатная структура" -Id $subdiv.OssSubdivisionID).FieldValues.VitroOrgLogin.LookupId -eq $subdiv.GroupID){
    Write-Host "У подразделения" $subdiv.Subdivision "уже проставлена группа." -ForegroundColor Green
    }
    else{
    Set-PnPListItem -List "Организационно-штатная структура" -ContentType "Подразделение" -Identity $subdiv.OssSubdivisionID -Values $GroupValues
}
}