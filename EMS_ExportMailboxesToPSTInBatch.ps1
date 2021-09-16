# 1. Create a txt file with mailboxes to be exported (1 mailbox alias per line).
# 2. Run following commands to create the mailbox batch:

$Mailboxes  = Get-Content c:\temp\safelinkmailboxes.txt

ForEach ($Mailbox in $Mailboxes) {
    New-MailboxExportRequest -Mailbox $Mailbox -FilePath "\\server\share\($Mailbox).pst" -BatchName Job1
}
    
# 3. Check batch move status with this command:

Get-MailboxExportRequest –BatchName Job1 | Get-MailboxExportRequestStatistics | `
Select DisplayName,StartTimeStamp,CompletionTimeStamp,TotalMailboxSize,Status
