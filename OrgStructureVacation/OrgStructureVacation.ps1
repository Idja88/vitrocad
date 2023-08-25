param (
    [Parameter(Mandatory=$true)]
    [string]$url,
    [string]$orglistname = "Организационно-штатная структура",
    [string]$empcontent = "0x010040D9D13AF3634AACB514CB30B54F2CAE004BC5116E5FCC734388387FAF0125CB7D"
)

function Set-Substitute($listName, $itemId, $substituteDelayed) {
    $com = ""
    foreach ($x in $substituteDelayed) {
        $com += $x.LookupId.ToString() + ";#" + $x.LookupValue + ";#"
    }
    Set-PnPListItem -List $listName -Identity $itemId -Values @{"VitroOrgSubstitute" = $com}
}

function Reset-SubstituteAndVacationInfo($listName, $itemId) {
    Set-PnPListItem -List $listName -Identity $itemId -Values @{"VitroOrgSubstitute" = $null; "VitroOrgSubstituteDelayed" = $null; "VitroOrgVacationStartDate" = $null; "VitroOrgVacationDueDate" = $null}
}

$admin = Get-PnpStoredCredential -Name $url -Type PSCredential
Connect-PnPOnline -Url $url -Credentials $admin
$dt = Get-Date
$emps = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true -and $_.FieldValues.VitroOrgSubstituteDelayed -ne $null}

foreach($i in $emps){
    if($null -ne $i.FieldValues.VitroOrgVacationStartDate -and $dt.Date -eq $i.FieldValues.VitroOrgVacationStartDate.AddDays(1).Date){
        Set-Substitute -listName $orglistname -itemId $i.FieldValues.ID -substituteDelayed $i.FieldValues.VitroOrgSubstituteDelayed
    }
    ElseIf($null -ne $i.FieldValues.VitroOrgVacationStartDate -and $dt.Date -ge $i.FieldValues.VitroOrgVacationStartDate.AddDays(1).Date -and $null -eq $i.FieldValues.VitroOrgSubstitute){
        Set-Substitute -listName $orglistname -itemId $i.FieldValues.ID -substituteDelayed $i.FieldValues.VitroOrgSubstituteDelayed
    }
    ElseIf($null -ne $i.FieldValues.VitroOrgVacationDueDate -and $dt.Date -gt $i.FieldValues.VitroOrgVacationDueDate.AddDays(1).Date){
        Reset-SubstituteAndVacationInfo -listName $orglistname -itemId $i.FieldValues.ID
    }
}