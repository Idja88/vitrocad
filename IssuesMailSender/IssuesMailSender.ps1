#Объявляем переменные
param (
    [string]$url,
    [string]$Design = $($url + "/Design"),
    [string]$issuelistname = "Замечания",
    [string]$orglistname = "Организационно-штатная структура",
    [string]$fizlistname = "Физические лица",
    [string]$empcontent = "0x010040D9D13AF3634AACB514CB30B54F2CAE004BC5116E5FCC734388387FAF0125CB7D",
    [string]$dt = $(Get-Date -Format "dd/MM/yyyy"),
    [string]$smtpserver,
    [string]$from,
    [string]$subject = "Сроки выполнения замечаний в Vitro-CAD [$($dt)]"
)

function Add-Table()
{
    param (
    [Parameter(Mandatory=$true,Position=0)]
    $issues
    )

    $HtmlTable = "<table border='1' align='Left' cellpadding='2' cellspacing='0' style='color:black;font-family:arial,helvetica,sans-serif;text-align:left;'>
    <tr style ='font-size:13px;font-weight: normal;background: #FFFFFF'>
    <th align=center><b>ID</b></th>
    <th align=center><b>Проект</b></th>
    <th align=center><b>Замечание</b></th>
    <th align=center><b>Срок</b></th>
    <th align=center><b>Автор</b></th>
    <th align=center><b>Статус</b></th>
    <th align=center><b>Файл</b></th>
    </tr>"

    foreach ($row in $issues){

            #Форматируем данные
            $issueId = $row.FieldValues.ID
            $fileId = $row.FieldValues.VitroBaseLibraryItemUniqueId

            $duedate = if($null -ne $row.FieldValues.VitroBaseCommentDate){$row.FieldValues.VitroBaseCommentDate.ToString('d')}else{$row.FieldValues.VitroBaseCommentDate}

            $issuelink = "<a href='$Design/_layouts/15/Vitro/TableView/ListView.aspx?List=CommentList&listname=$issuelistname&ID=$issueId'>" + $row.FieldValues.ID + "</a>"
            $filelink = "<a href='$url/_layouts/15/Vitro/ProtocolHandler/VitroProtocolHandler.aspx?target=vitro://vitro/Design{$fileId}'>" + $row.FieldValues.VitroBaseLibraryItemName + "</a>"
        
            #Заполняем таблицу
            $HtmlTable += "<tr style='font-size:13px;background-color:#FFFFFF'>
            <td>" + $issuelink + "</td>
            <td>" + $row.FieldValues.VitroBaseCommentProject + "</td>
            <td>" + $row.FieldValues.VitroBaseCommentNote + "</td>
            <td>" + $duedate + "</td>
            <td>" + $row.FieldValues.VitroBaseCommentAuthor.Lookupvalue + "</td>
            <td>" + $row.FieldValues.VitroBaseCommentStatus.Lookupvalue + "</td>
            <td>" + $filelink + "</td>
            </tr>"
            }

    Return $HtmlTable += "</table>"
}

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

Connect-PnPOnline $url -CurrentCredentials

#собираем текущих пользователей ОШС
$emps = Get-PnPListItem -List $orglistname | Where-Object {$_.FieldValues.ContentTypeId -like $empcontent -and $_.FieldValues.VitroOrgDisplayInStructure -eq $true}

#Собираем почтовые адреса из списка физ. лиц
foreach ($emp in $emps){

    $PersonEmail = (Get-PnPListItem -List $fizlistname | Where-Object {$_.FieldValues.ID -eq $emp.FieldValues.VitroOrgPerson.LookupId}).FieldValues.Email

    $emp.FieldValues | Add-Member -MemberType NoteProperty -Name Mail -Value $PersonEmail -Force
}

Connect-PnPOnline $Design -CurrentCredentials

foreach ($emp in $emps)
{
    #Заголовки тела письма
    $open = "<br/>Мои замечания (Открытые):<br/>"
    $closed = "<br/>Замечания от меня (Выполненные):<br/>"

    #Собираем открытые поручения
    $openissues = Get-PnPListItem -List $issuelistname | Where-Object {$_.FieldValues.VitroBaseCommentStatus.LookupId -in (1, 2, 4) -and $_.FieldValues.VitroBaseCommentAssignTo.LookupId -eq $emp.FieldValues.ID} | Sort-Object -Property {$_.FieldValues.VitroBaseCommentProject}

    #Собираем закрытые поручения
    $closedissues = Get-PnPListItem -List $issuelistname | Where-Object {$_.FieldValues.VitroBaseCommentStatus.LookupId -eq 3 -and $_.FieldValues.VitroBaseCommentAuthor.LookupId -eq $emp.FieldValues.ID} | Sort-Object -Property {$_.FieldValues.VitroBaseCommentProject}

    #Открытые поручения
    if($null -ne $openissues){$odata = Add-Table $openissues}else{$odata='';$open = ''}

    #Закрытые поручения
    if($null -ne $closedissues){$cdata = Add-Table $closedissues}else{$cdata='';$closed = ''}

    #Отправка почты
    $to = $emp.FieldValues.Mail
    $body = "Здравствуйте.<br/>Ниже перечислены замечания Vitro-CAD и сроки их исполнения:<br/>" + $open + $odata + $closed + $cdata
    if(($null -ne $openissues) -and ($null -ne $closedissues))
    {
        Send-Mail -To $to -body $body -From $from -Subject $subject -SMTPServer $smtpserver
    }
}