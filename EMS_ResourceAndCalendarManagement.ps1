# Use Add-MailboxPermission to add new permission, Set-MailboxPermission to modify existing & Remove-MailboxPermission to delete.
# Best practice is to use a Group Identity to grant permissions once and the add/remove the users from that Group.

# Script to manage Room Mailbox permissions:

$Rooms = Get-Mailbox | ? {($_.RecipientTypeDetails -eq "RoomMailbox") -and ($_.alias -like "*Room*")} 
Foreach ($Room in $Rooms) { 
$Calendar = $Room.alias +":\Calendar" 
Get-MailboxFolderPermission -Identity $Calendar -user UserOrGroupIdentity
}

# Script to grant Full Control permission to an resource mailbox:
$Rooms = Get-Mailbox | ? {($_.RecipientTypeDetails -eq "EquipmentMailbox") -and ($_.alias -like "training*")} 
Foreach ($Room in $Rooms) { 
Add-MailboxPermission -Identity $Room.alias -User UserOrGroupIdentity -AccessRights FullAccess -InheritanceType all
}

# Script to remove Full Control permission to an resource mailbox:
$Rooms = Get-Mailbox | ? {($_.RecipientTypeDetails -eq "RoomMailbox")}
Foreach ($Room in $Rooms) { 
Remove-MailboxPermission -Identity $Room.alias -User UserOrGroupIdentity -AccessRights FullAccess -InheritanceType All
}

# Script to grant/remove mailbox calendar permissions in bulk:
$TargetUser = "TargetUserName"
$Calendar = $TargetUser+":\Calendar" 
$UsersOrGroups = @("User1","Group1","Group2") 

Foreach ($UserOrGroup in $UsersOrGroups) { 
Add-MailboxFolderPermission -Identity $Calendar -User $UserOrGroup -AccessRights Reviewer 
}

<# 

-Roles and permissions granted are described in the following list:

* Author: CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
* Contributor: CreateItems, FolderVisible
* Editor: CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
* None: FolderVisible
* NonEditingAuthor: CreateItems, FolderVisible, ReadItems
* Owner: CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems
* PublishingEditor: CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems
* PublishingAuthor: CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems
* Reviewer: FolderVisible, ReadItems

- The following roles apply specifically to Calendar folders:

* AvailabilityOnly: View only availability data
* LimitedDetails: View availability data with subject and location

#>
