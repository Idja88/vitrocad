# Добавление Microsoft SharePoint Snap-in в текущую сессию
If($null -eq (Get-PsSnapin | Where-Object {$_.Name -eq "Microsoft.SharePoint.PowerShell"})){
    Write-Host -ForegroundColor White "Loading SharePoint Powershell Snapin"
    Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
}

#Объявлем переменные
$url = "http://vitro2.kazmintech.kz/"
$orglistname = "Организационно-штатная структура"
$fizlistname = "Физические лица"
$divcontent = "0x0100CE4B4034AA87410EB92561C8318E2C16005639E40A0D3F6342A19FF44EB45B5A85"
$empcontent = "0x010040D9D13AF3634AACB514CB30B54F2CAE004BC5116E5FCC734388387FAF0125CB7D"
$ProcessError = @()

$admin = Get-PnpStoredCredential -Name $url -Type PSCredential
Connect-PnPOnline $url -Credentials $admin

#Синхронизируем учётные записи из AD
$users = (Get-SPUser -Web $url).Id

foreach($user in $users){
    Set-SPUser -Web $url -Identity $user -SyncFromAD
}

#Собираем объекты Подразделений и Сотрудников
$emps = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true}
$divs = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $divcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true} | Select -Skip 1 #Убираем из массива корень структуры

foreach ($div in $divs) {
        
    #Проверка на существующие группы, как подразделения
    if($null -ne (Get-PnPGroup -Identity $div.FieldValues.Title)){
        Write-Host "Подразделение - " $div.FieldValues.Title "уже добавлено как группа пользователей." -ForegroundColor Green
        }
    else{
       $NewGrp = New-PnPGroup -Title $div.FieldValues.Title -Owner $admin.UserName -ErrorAction SilentlyContinue -ErrorVariable ProcessError
       if($ProcessError) {
        Write-Host "Ошибка с подразделением" $div.FieldValues.Title -ForegroundColor Red
        }
        else{
        Write-Host "Подразделение - " $div.FieldValues.Title "добавлено как группа пользователей." -ForegroundColor Yellow
       }
}
    #Забираем ID групп пользователей
    $divGrp = (Get-PnPGroup -Identity $div.FieldValues.Title).Id
    $div.FieldValues | Add-Member -MemberType NoteProperty -Name GroupId -Value $divGrp -Force

    [string]$strGrp = $div.FieldValues.GroupId
    $GroupValues =@{"VitroOrgLogin" = $strGrp}

    #Проверка на уже проставленные группы
    if((Get-PnPListItem -List $orglistname -Id $div.Id).FieldValues.VitroOrgLogin.LookupId -eq $div.FieldValues.GroupID){
    Write-Host "У подразделения" $div.FieldValues.Title "уже проставлена аналогичная группа пользователей." -ForegroundColor Green
    }
    else{
    $SetGrp = Set-PnPListItem -List $orglistname -ContentType "Подразделение" -Identity $div.Id -Values $GroupValues -ErrorAction SilentlyContinue -ErrorVariable ProcessError
    if($ProcessError) {
        Write-Host "Ошибка с подразделением" $div.FieldValues.Title -ForegroundColor Red
        }
        else {
        Write-Host "К подразделению" $div.FieldValues.Title "добавлена аналогичная группа пользователей."-ForegroundColor Yellow
        }
    }
}

foreach ($emp in $emps){
    $LoginName = (Get-PnPListItem -List $fizlistname | Where-Object {$_.FieldValues.ID -eq $emp.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupValue
    $LoginId = (Get-PnPListItem -List $fizlistname | Where-Object {$_.FieldValues.ID -eq $emp.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupId

    $emp.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLogin -Value $LoginName -Force
    $emp.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLoginId -Value $LoginId -Force

    if(Get-PnPGroupMembers -Identity $emp.FieldValues.VitroOrgParentId.LookupValue | Where-Object {$_.Id -eq $emp.FieldValues.PersonLoginId}){
    Write-Host "Пользователь - " $emp.FieldValues.Title "уже находится в группе." -ForegroundColor Green
    }
    else {
    $UserToGrp = Add-PnPUserToGroup -Identity $emp.FieldValues.VitroOrgParentId.LookupValue -LoginName $emp.FieldValues.PersonLogin -ErrorAction SilentlyContinue -ErrorVariable ProcessError
    if($ProcessError) {
    Write-Host "Невозможно найти группу" -ForegroundColor Red
    }
    else {
    Write-Host "Пользователь - " $emp.FieldValues.Title "добавлен в группу." -ForegroundColor Yellow
    }
  }
}

#Проверка на соответствие Сотрудников в Группе Пользователей и удаление лишних (перевод сотрудника в др. отдел)
foreach ($div in $divs){

$emparr = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true -and $_.FieldValues.VitroOrgParentId.LookupValue -eq $div.FieldValues.Title}

foreach ($i in $emparr){
    $EmpLoginName = (Get-PnPListItem -List $fizlistname | Where-Object {$_.FieldValues.ID -eq $i.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupValue
    $EmpLoginId = (Get-PnPListItem -List $fizlistname | Where-Object {$_.FieldValues.ID -eq $i.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupId

    $i.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLogin -Value $EmpLoginName -Force
    $i.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLoginId -Value $EmpLoginId -Force
}

$grparr = Get-PnPGroupMembers -Identity $div.FieldValues.GroupId | 
    ForEach-Object {
        if($_.Id -in $emparr.FieldValues.PersonLoginId) { 
            Write-Host "Сотрудник" $_.Title "принадлежит данному подразделению." -ForeGroundColor Green} 
        else {
            #Убираем логин пользователя, который больше не является сотрудником подразделения
            Write-Host "Сотрудник" $_.Title "больше не принадлежит данному подразделению, и будет удалён из группы." -ForegroundColor Yellow
            $Rmvusr = Remove-PnPUserFromGroup -LoginName $_.LoginName -Identity $div.FieldValues.GroupId
        }
        }
}

$ProcessError.Count