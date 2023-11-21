param (
    [Parameter(Mandatory=$true)]
    [string]$url,
    [string]$orglistname = "Организационно-штатная структура",
    [string]$fizlistname = "Физические лица",
    [string]$divcontent = "0x0100CE4B4034AA87410EB92561C8318E2C16005639E40A0D3F6342A19FF44EB45B5A85",
    [string]$empcontent = "0x010040D9D13AF3634AACB514CB30B54F2CAE004BC5116E5FCC734388387FAF0125CB7D",
    [string]$funccontent = "0x0100014407A2350F49DD996BFE060D9B018700F2B87C61FA93FB42AAA238DAB6F082C6"
)

function Test-ModuleInstalled {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )
    return $null -ne (Get-Module -ListAvailable -Name $ModuleName)
}

function Add-LoginProperties {
    param (
        $Item,
        $ListItems
    )
    $LoginName = ($ListItems | Where-Object {$_.FieldValues.ID -eq $Item.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupValue
    $LoginId = ($ListItems | Where-Object {$_.FieldValues.ID -eq $Item.FieldValues.VitroOrgPerson.LookupId}).FieldValues.VitroOrgLogin.LookupId

    $Item.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLogin -Value $LoginName -Force
    $Item.FieldValues | Add-Member -MemberType NoteProperty -Name PersonLoginId -Value $LoginId -Force
}

# Добавление MicrosoftSharePointPowershell Module в текущую сессию
if (-not (Test-ModuleInstalled -ModuleName "Microsoft.SharePoint.PowerShell")) {
    Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
}

$admin = Get-PnpStoredCredential -Name $url -Type PSCredential
Connect-PnPOnline $url -Credentials $admin

#Собираем объекты Подразделений и Сотрудников
$users = Get-SPUser -Web $url
$orgitems = Get-PnPListItem -List $orglistname
$fizitems = Get-PnPListItem -List $fizlistname
$emps = $orgitems | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true}
$divs = $orgitems | Where-Object {$_.FieldValues.ContentTypeId -like $divcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true} | Select-Object -Skip 1 #Убираем из массива корень структуры
$funcs = $orgitems | Where-Object {$_.FieldValues.ContentTypeId -like $funccontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true}
#Собираем список физ лиц без заполненной эл почты, при необходимости синхронизации данных из AD
#$pers = $fizitems | Where-Object {$null -ne $_.FieldValues.VitroOrgLogin}

#Синхронизируем данные учётных записей из AD
foreach($user in $users){
    Set-SPUser -Web $url -Identity $user.ID -SyncFromAD
}

#Обновляем поля Телефон и Эл.Почта, при необходимости синхронизации данных из AD
<#
foreach ($per in $pers){
    $identity = $users | Where-Object {$_.ID -eq $per.FieldValues.VitroOrgLogin.LookupId}
  
    $aduser = Get-ADUser -Identity $identity.UserId.NameId -Properties "mail","telephoneNumber"
    
    if($null -ne $aduser.mail){
      if($null -eq $per.FieldValues.Email -or $aduser.mail -ne $per.FieldValues.Email){
      Set-PnPListItem -List $fizlistname -Identity $per.FieldValues.ID -Values @{"Email" = $aduser.mail}
      }
    }
    
    if($null -ne $aduser.telephoneNumber){
      if($null -eq $per.FieldValues.VitroOrgPhone -or $aduser.telephoneNumber -ne $per.FieldValues.VitroOrgPhone){
      Set-PnPListItem -List $fizlistname -Identity $per.FieldValues.ID -Values @{"VitroOrgPhone" = $aduser.telephoneNumber}
      }
    }
  }
#>

foreach ($div in $divs) {
    #Проверка на существующие группы, как подразделения и добавление их как параметр
    $divGrp = Get-PnPGroup -Identity $div.FieldValues.Title
    if($null -eq $divGrp) {
        $divGrpNew = New-PnPGroup -Title $div.FieldValues.Title -Owner $admin.UserName
        $div.FieldValues | Add-Member -MemberType NoteProperty -Name GroupId -Value $divGrpNew.Id -Force
        [string]$strGrp = $div.FieldValues.GroupId
        $GroupValues = @{"VitroOrgLogin" = $strGrp}
    }
    else {
        $div.FieldValues | Add-Member -MemberType NoteProperty -Name GroupId -Value $divGrp.Id -Force
        [string]$strGrp = $div.FieldValues.GroupId
        $GroupValues = @{"VitroOrgLogin" = $strGrp}
    }

    #Проверка на уже проставленные группы
    if($div.FieldValues.VitroOrgLogin.LookupId -ne $div.FieldValues.GroupID -or $null -eq $div.FieldValues.VitroOrgLogin.LookupId){
        Set-PnPListItem -List $orglistname -ContentType "Подразделение" -Identity $div.Id -Values $GroupValues
    }

    #Проверка на соответствие Сотрудников в Группе Пользователей и удаление лишних (перевод сотрудника в др. отдел)
    $divemps = $emps | Where-Object {$_.FieldValues.VitroOrgParentId.LookupId -eq $div.FieldValues.ID}
    $grpmbrs = Get-PnPGroupMembers -Identity $div.FieldValues.GroupId

    #Добавляем сотрудника в группу пользователей
    foreach ($divemp in $divemps){
        Add-LoginProperties -Item $divemp -ListItems $fizitems
        if($divemp.FieldValues.PersonLoginId -notin $grpmbrs.ID){
            Add-PnPUserToGroup -Identity $div.FieldValues.GroupId -LoginName $divemp.FieldValues.PersonLogin
        }
    }

    #Убираем логин пользователя, который больше не является сотрудником подразделения
    foreach($grpmbr in $grpmbrs){
        if($grpmbr.Id -notin $divemps.FieldValues.PersonLoginId){
            Remove-PnPUserFromGroup -LoginName $grpmbr.LoginName -Identity $div.FieldValues.GroupId
        }
    }
}

foreach ($func in $funcs) {
    #Проверка на существующие группы, как ф.группы и добавление их как параметр
    $funcGrp = Get-PnPGroup -Identity $func.FieldValues.Title
    if($null -eq $funcGrp) {
        $funcGrpNew = New-PnPGroup -Title $func.FieldValues.Title -Owner $admin.UserName
        $func.FieldValues | Add-Member -MemberType NoteProperty -Name GroupId -Value $funcGrpNew.Id -Force
        [string]$strFuncGrp = $func.FieldValues.GroupId
        $FuncGroupValues = @{"VitroOrgLogin" = $strFuncGrp}
    }
    else {
        $func.FieldValues | Add-Member -MemberType NoteProperty -Name GroupId -Value $funcGrp.Id -Force
        [string]$strFuncGrp = $func.FieldValues.GroupId
        $FuncGroupValues = @{"VitroOrgLogin" = $strFuncGrp}
    }

    #Проверка на уже проставленные группы
    if($func.FieldValues.VitroOrgLogin.LookupId -ne $func.FieldValues.GroupID -or $null -eq $func.FieldValues.VitroOrgLogin.LookupId){
        Set-PnPListItem -List $orglistname -ContentType "Функциональная группа" -Identity $func.Id -Values $FuncGroupValues
    }

    #Проверка на соответствие Сотрудников в Группе Пользователей и удаление лишних (перевод сотрудника в др. отдел)
    $funcemps = $emps | Where-Object {$_.FieldValues.ID -in $func.FieldValues.VitroOrgMemberList.LookupId}
    $funcgrpmbrs = Get-PnPGroupMembers -Identity $func.FieldValues.GroupId

    #Добавляем сотрудника в группу пользователей
    foreach ($funcemp in $funcemps){
        Add-LoginProperties -Item $funcemp -ListItems $fizitems
        if($funcemp.FieldValues.PersonLoginId -notin $funcgrpmbrs.ID){
            Add-PnPUserToGroup -Identity $func.FieldValues.GroupId -LoginName $funcemp.FieldValues.PersonLogin
        }
    }

    #Убираем логин пользователя, который больше не является сотрудником ф.группы
    foreach($funcgrpmbr in $funcgrpmbrs){
        if($funcgrpmbr.Id -notin $funcemps.FieldValues.PersonLoginId){
            Remove-PnPUserFromGroup -LoginName $funcgrpmbr.LoginName -Identity $func.FieldValues.GroupId
        }
    }
}