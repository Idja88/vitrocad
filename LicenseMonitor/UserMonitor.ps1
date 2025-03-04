$vitroserver = ""

$counter1 = '\\{0}\Веб-служба(Sharepoint - 80)\Количество текущих подключений' -f $vitroserver
$counter2 = '\\{0}\Веб-служба(Sharepoint - 80)\Максимальное количество подключений' -f $vitroserver
$counter3 = '\\{0}\Веб-служба(Sharepoint - 80)\Количество подключенных анонимных пользователей' -f $vitroserver
$counter4 = '\\{0}\Веб-служба(Sharepoint - 80)\Максимальное количество анонимных пользователей' -f $vitroserver
$counter5 = '\\{0}\Веб-служба(Sharepoint - 80)\Количество подключенных неанонимных пользователей' -f $vitroserver
$counter6 = '\\{0}\Веб-служба(Sharepoint - 80)\Максимальное количество неанонимных пользователей' -f $vitroserver

$Counters = @(
    $counter1,
    $counter2,
    $counter3,
    $counter4,
    $counter5,
    $counter6
)

$OutputFile = "UserCounter.csv"

Get-Counter -Counter $Counters | ForEach {
    $_.CounterSamples | ForEach {
        [pscustomobject]@{
            TimeStamp = $_.TimeStamp
            Path = $_.Path
            Value = $_.CookedValue
        }
    }
} | Export-Csv -Path $OutputFile -Append -NoTypeInformation -Encoding UTF8