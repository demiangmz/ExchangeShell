# Delete a large item folder from mailbox using EWS 

$MailboxName = 'Name@Mailbox'
$dllpath = ".\Microsoft.Exchange.WebServices.dll"

[void][Reflection.Assembly]::LoadFile($dllpath)
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)

# If command times out add this parameter
# Service.Timeout= X
# where X is the timeout in milliseconds

$Service.AutodiscoverUrl($MailboxName,{$true})
$RootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)

$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RootFolderID)
$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$Response = $RootFolder.FindFolders($FolderView)


# Change method to HardDelete if you want to delete all the items permanently

ForEach ($Folder in $Response.Folders) {
  if($folder.DisplayName -eq "AName") {
    $folder.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
    } 
 }
