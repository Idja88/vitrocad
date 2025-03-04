$vitroserver = ""

$counter1 = '\\{0}\Веб-служба(VitroLicenseService)\Количество текущих подключений' -f $vitroserver
$counter2 = '\\{0}\Веб-служба(VitroLicenseService)\Максимальное количество подключений' -f $vitroserver

$Counters = @(
    $counter1,
    $counter2
)
$OutputFile = "LicenseCounter.csv"

Get-Counter -Counter $Counters | ForEach {
    $_.CounterSamples | ForEach {
        [pscustomobject]@{
            TimeStamp = $_.TimeStamp
            Path = $_.Path
            Value = $_.CookedValue
        }
    }
} | Export-Csv -Path $OutputFile -Append -NoTypeInformation -Encoding UTF8