# 1. Create a txt file with mailboxes to be moved (1 mailbox alias per line).
# 2. Run following commands to create the mailbox batch:
$Mailboxes  = Get-Content c:\temp\mailboxes.txt
	
ForEach ($Mailbox in $Mailboxes) {
  New-MoveRequest -Identity $Mailbox -TargetDatabase DB -BatchName 1stBatch -BadItemLimit 25
}

# 3. Check batch move status with this command:
Get-MoveRequest -BatchName  1st Batch | Get-MoveRequestStatistics | Select DisplayName,StartTimeStamp,CompletionTimeStamp,TotalMailboxSize,Status

# 4. Check batch overall move time:
Get-MoveRequest -BatchName 1stBatch | Get-MoveRequestStatistics | Select overallduration | Sort-Object overallduration -Descending
