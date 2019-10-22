add-pssnapin *exchange* -erroraction SilentlyContinue

$i = 0    

$Mailboxes = Get-Mailbox

$Results = foreach($Mailbox in $Mailboxes)
{
    $Folders = $MailBox |
        Get-MailboxFolderStatistics |
        Measure-Object |
        Select-Object -ExpandProperty Count

    New-Object -TypeName PSCustomObject -Property @{
        Username    = $Mailbox.Alias
        FolderCount = $Folders
        DisplayName = $Mailbox.DisplayName
        }
    

   $i++
   Write-Progress -activity "Checking mailboxes" -status "Checked so far: $i of $($Mailboxes.Count)" -percentComplete (($i / $Mailboxes.Count)  * 100)
}


$Results | Select-Object -Property Username, DisplayName, FolderCount | convertto-csv -NoTypeInformation | % {$_.Replace('"','')} | out-file .\FolderCountExport.csv -fo -en ascii
