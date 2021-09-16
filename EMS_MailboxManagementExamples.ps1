# Search Mailbox for large items
Get-Mailbox -ResultSize Unlimited | Get-MailboxFolderStatistics -IncludeAnalysis -FolderScope All | `
Where-Object {(($_.TopSubjectSize -Match "MB") -and ($_.TopSubjectSize -GE 50.0)) -or ($_.TopSubjectSize -Match "GB")} | `
Select-Object Identity, TopSubject, TopSubjectSize | `
Export-CSV -path "C:\report.csv" -notype

# Delete specific message from ALL mailboxes
# First search for the messages and import them to your mailbox to check the deletion won't affect valid email:
Get-ExchangeServer | where {$_.isMailboxServer -eq $true} | Get-Mailbox | `
Search-Mailbox -SearchQuery {Subject:"FW:Consult" AND From:"sender@domain.com" AND To:"Finance@company.com" AND Received:today/yesterday/received:>12/31/2021 AND received:<1/1/2022} -targetmailbox "targetMailbox" -targetfolder "Test" -logonly -loglevel full
 
#Then run the same command with the -DeleteContent parameter set:
Get-ExchangeServer | where {$_.isMailboxServer -eq $true} | Get-Mailbox | `
Search-Mailbox -SearchQuery {Subject:"FW:Consult" AND From:"sender@domain.com" AND To:"recipient@company.com"} -LogLevel Full -DeleteContent

#Empty Deleted Items folder from mailbox
Search-Mailbox -Identity username -SearchQuery '#deleted items#' -DeleteContent

#Delete Dumpster items from mailbox
Search-Mailbox "Mailbox.Name" -SearchDumpsterOnly -DeleteContent

#Find Mailboxes with forwarding addresses
Get-mailbox | select DisplayName,ForwardingAddress | where {$_.ForwardingAddress -ne $null}
