param (
    [Parameter(Mandatory=$true)]
    [string]$url,
    [string]$contractlistname = "Реестр договоров",
    [string]$from,
    [string]$smtpserver,
    [int]$end = 15
)

Function Send-Mail
{   [CmdletBinding()]
    Param (
        [string]$To,
        [string]$body,
        [string]$From,
        [string]$Subject,
        [string]$SMTPServer,
        [string]$Priority = "High"
    )
    
    Do {
        Try {
            Send-MailMessage -smtpserver $SMTPServer -from $from -to $To -subject $Subject -body $body -bodyashtml -Priority $Priority -Encoding UTF8 -ErrorAction Stop
            $Exit = 4
        }
        Catch {
            $Exit ++
            Write-Verbose "Failed to send message because: $($Error[0])"
            Write-Verbose "Try #: $Exit"
            If ($Exit -eq 4)
            {   Write-Warning "Unable to send message!" $To
            }
        }
    } Until ($Exit -eq 4)
}
#main data
Connect-PnPOnline $url -CurrentCredentials

#collect contracts
$dt = Get-Date
$cons = Get-PnPListItem -List $contractlistname | Where-Object {$_.FieldValues.VitroBaseContractFinishDate -ne $null}
foreach($con in $cons){
    $ts = New-TimeSpan -Start $dt.Date -End $con.FieldValues.VitroBaseContractFinishDate.AddDays(1).Date
    if($ts.Days -eq $end){
        $to = $con.FieldValues.ContractHolderMail
        $subject = "Vitro. Договор " + $con.FieldValues.VitroBaseNumber + " заканчивается."
        $ID = $con.FieldValues.ID
        $body = "<span>Уважаемый(ая) " + $con.FieldValues.ContractHolderName + ", до конца договора под номером " + "<b><a href='$Main/Lists/ContractList/DispForm.aspx?ID=$ID'>" + $con.FieldValues.VitroBaseNumber + "</a></b>" + " осталось 15 дней.</span>"

        Send-Mail -To $to -body $body -From $from -Subject $subject -SMTPServer $smtpserver
    }
}