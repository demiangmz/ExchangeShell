# Script to track email messages on all servers for a specific Recipient
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true} |  `
Get-MessageTrackingLog -Start (get-date).addhours(-12) -Recipients "recipíent@domain.com" | `
Select-Object Timestamp,ServerHostname,ClientHostname,ClientIp,Source,SourceContext,EventId,Sender,Recipients,TotalBytes | `
Sort-Object -property Timestamp

# Another example with Sender
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true} | `
Get-MessageTrackingLog -Start (Get-Date).AddDays(-1) -Sender "sender@domain.com" -MessageSubject "RE: February Reports" | `
Select-Object Timestamp,ServerHostname,ClientHostname,ClientIp,Source,SourceContext,EventId,Sender,Recipients,TotalBytes | `
Sort-Object -property Timestamp

# Example when you don't know the Sender but you know the Domain:
Get-ExchangeServer | where {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true} | `
Get-MessageTrackingLog -Start (get-date).adddays(-1) -Recipients "recipient@domain.com" -ResultSize Unlimited | `
Where {$_.Sender -like "sender@domain.com"} | `
Select-Object Timestamp,ServerHostname,ClientHostname,ClientIp,Source,SourceContext,EventId,Sender,Recipients,TotalBytes,MessageSubject | `
Sort-Object -property Timestamp

# Get number of messages with Failed status from last 7 days:
$7DaysOfTracking = Get-ExchangeServer | Get-MessageTrackingLog -EventID Fail -Start (get-date).adddays(-1) -ResultSize unlimited
$7DaysOfTracking | Measure-Object
$7DaysOfTracking | Select-object Timestamp,ClientIP,ClientHostname,ServerIP,ServerHostname,SourceContext,ConnectorID,Source,EventId,MessageId,@{label="Recipients";expression={[string]($_.Recipients | foreach {$_.tostring().split("/")[-1]})}},@{label="RecipientStatus";expression={[string]($_.RecipientStatus | foreach {$_.tostring().split("/")[-1]})}},RecipientCount,MessageSubject,Sender | Export-Csv -Path MessageTracking7Days.csv

#Convert "Recipients" and other string arrays to readable fields:
$MessageTrackingResultVariable | `
Select Timestamp,ClientIP,ClientHostname,ServerIP,ServerHostname,SourceContext,ConnectorID,Source,EventId,MessageId,@{label="Recipients";expression={[string]($_.Recipients | `
foreach {$_.tostring().split("/")[-1]})}},@{label="RecipientStatus";expression={[string]($_.RecipientStatus | `
foreach {$_.tostring().split("/")[-1]})}},RecipientCount,MessageSubject,Sender | `
Export-Csv -Path MessageTracking7Days.csv

#Check message latencies in the past hour:
Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddHours(-1) -EventID SEND | `
where {(($_.MessageLatency).TotalSeconds -gt 90)} | `
Select Timestamp,ClientIP,ClientHostname,ServerIP,ServerHostname,SourceContext,ConnectorID,Source,EventId,MessageId,@{label="Recipients";expression={[string]($_.Recipients | `
foreach {$_.tostring().split("/")[-1]})}},@{label="RecipientStatus";expression={[string]($_.RecipientStatus | `
foreach {$_.tostring().split("/")[-1]})}},RecipientCount,MessageSubject,Sender, @{Label="LatencyMil"; Expression={$($_.MessageLatency).TotalMilliseconds}} | `
FT TimeStamp, SourceContext,Recipients, Sender, LatencyMil

#If you are just interested in a particular day and also want to know which HUB server delivered it each message:
Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddHours(-1) -EventID DELIVER | `
Select TimeStamp, ClientHostname, @{Label="LatencyMil"; Expression={$($_.MessageMeLatency).TotalMilliseconds}} | `
Export-Csv Latency.csv -NoTypeInformation

#Check how many emails a distribution list from your domain has received in the last month
Get-TransportServer | Get-MessageTrackingLog -start (Get-Date).AddDays(-30) -EventId "expand" -ResultSize Unlimited | `
where {($_.RelatedRecipientAddress -like "emailaddress@domainX.com") -and ($_.Sender -notmatch "domainX.com")} | `
Measure-Object 
