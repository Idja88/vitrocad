#Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteURL,
    [string]$ListName = "Проекты",
    [int]$VersionsToKeep = 10
)

Connect-PnPOnline -Url $SiteURL -CurrentCredentials
 
$List = Get-PnPList -Identity $ListName
  
$Ctx= Get-PnPContext
 
$global:counter=0
$ListItems = Get-PnPListItem -List $ListName -Fields FileLeafRef -PageSize 2000 -ScriptBlock { Param($items) $global:counter += $items.Count; Write-Progress `
                -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting Files of '$($List.Title)'" `
                    -Status "Processing Files $global:Counter of $($List.ItemCount)";} | Where-Object {($_.FileSystemObjectType -eq "File")}
Write-Progress -Activity "Completed Retrieving Files!" -Completed
 
$TotalFiles = $ListItems.count
$Counter = 1
ForEach ($Item in $ListItems)
{
    $File = $Item.File
    $Versions = $File.Versions
    $Ctx.Load($File)
    $Ctx.Load($Versions)
    $Ctx.ExecuteQuery()
  
    Write-host -f Yellow "Scanning File ($Counter of $TotalFiles):"$Item.FieldValues.FileRef
    $VersionsCount = $Versions.Count
    $VersionsToDelete = $VersionsCount - $VersionsToKeep
    If($VersionsToDelete -gt 0)
    {
        write-host -f Cyan "`t Total Number of Versions of the File:" $VersionsCount
        For($i=0; $i -lt $VersionsToDelete; $i++)
        {
            write-host -f Cyan "`t Deleting Version:" $Versions[0].VersionLabel
            $Versions[0].DeleteObject()
        }
        $Ctx.ExecuteQuery()
        Write-Host -f Green "`t Version History is cleaned for the File:"$File.Name
    }
    $Counter++
}
