$url = ""
$orglistname = "Организационно-штатная структура"
$empcontent = "0x010040D9D13AF3634AACB514CB30B54F2CAE004BC5116E5FCC734388387FAF0125CB7D"

Connect-PnPOnline -Url $url -CurrentCredentials
$dt = Get-Date #-Format "dd.MM.yyyy"

$emps = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true -and $_.FieldValues.VitroOrgSubstituteDelayed -ne $null}

foreach($i in $emps){
   if($null -ne $i.FieldValues.VitroOrgVacationStartDate -and $dt.Date -eq $i.FieldValues.VitroOrgVacationStartDate.AddDays(1).Date){
      switch ($i.FieldValues.VitroOrgSubstituteDelayed.LookupId.Count){
        #{$_ -ge 2} {$com = $i.FieldValues.VitroOrgSubstituteDelayed | ForEach-Object {$_.LookupId.ToString() + ";#" + $_.LookupValue.ToString() + ";#"}
        #Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $com}
        {$_ -ge 2}{
        $com = ""
        foreach($x in $i.FieldValues.VitroOrgSubstituteDelayed){$com += $x.LookupId.ToString() + ";#" + $x.LookupValue +";#"}
        Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $com}
        }
        Default {Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $i.FieldValues.VitroOrgSubstituteDelayed.LookupId}}
      }
   }
   Elseif($null -ne $i.FieldValues.VitroOrgVacationStartDate -and $dt.Date -ge $i.FieldValues.VitroOrgVacationStartDate.AddDays(1).Date -and $i.FieldValues.VitroOrgSubstitute -eq $null){
      switch ($i.FieldValues.VitroOrgSubstituteDelayed.LookupId.Count){
       #{$_ -ge 2} {$com = $i.FieldValues.VitroOrgSubstituteDelayed | ForEach-Object {$_.LookupId.ToString() + ";#" + $_.LookupValue.ToString()+ ";#"}
       #Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $com}
       {$_ -ge 2}{
       $com = ""
       foreach($x in $i.FieldValues.VitroOrgSubstituteDelayed){$com += $x.LookupId.ToString() + ";#" + $x.LookupValue +";#"}
       Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $com}
       }
       Default {Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $i.FieldValues.VitroOrgSubstituteDelayed.LookupId}}
     }
   }
   Elseif($null -ne $i.FieldValues.VitroOrgVacationDueDate -and $dt.Date -gt $i.FieldValues.VitroOrgVacationDueDate.AddDays(1).Date){
     Set-PnPListItem -List $orglistname -Identity $i.FieldValues.ID -Values @{"VitroOrgSubstitute" = $null; "VitroOrgSubstituteDelayed" = $null; "VitroOrgVacationStartDate" = $null; "VitroOrgVacationDueDate" = $null}
   }
}