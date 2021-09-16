# Convert Mailbox Object to Contact Object retaining legacyExchangDN attribute as x500 address

$DomainController = (Get-ADServerSettings).DefaultConfigurationDomainController.domain
 
$MailboxList = Get-Content -Path C:\temp\Mailboxes.txt
 
ForEach ($Mbx in $MailboxList) {
	
	$Mailbox = Get-Mailbox $Mbx
	$EmailAddresses = $Mailbox.EmailAddresses
	$EmailAddresses += "x500:$($Mailbox.LegacyExchangeDN)"

	Disable-Mailbox -Id $mailbox.Identity -Confirm:$False -DomainController $DomainController
	Start-Sleep -Seconds 30
	$smtp = $mailbox.primarysmtpaddress.local + "@domain.com"
	$contact = New-MailContact -alias $mailbox.alias -Name $mailbox.name -ExternalEmailAddress $smtp `
		-OrganizationalUnit "OU=ExternalUsers,DC=domain,DC=com" -DomainController $DomainController
	
	Set-MailContact -Id $contact -EmailAddresses $EmailAddresses -EmailAddressPolicyEnabled $False -DomainController $DomainController
}
