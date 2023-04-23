# PowerPhish
This PowerShell script opens a realistic looking window that asks for credentials. If the computer is part of a domain the credentials are validated. Afterwards, the loot is uploaded to Pastebin as a private paste. If the entered credentials are empty or wrong, the script will prompt the user again. This script was created with a "Bad USB" / HID / Ducky attack in mind. 

![Screenshot](/Screenshots/window_english.PNG)

# Usage
You must set the following variables for the script to upload your loot to Pastebin:
- `$pastebin_username` is your Pastebin username
- `$pastebin_password` is your Pastebin password
- `$pastebin_dev_key` is your Pastebin API key (which can be found at https://pastebin.com/doc_api)

You can change the following variables:
- `$target` sets the "connection" name which is displayed in the message. (In the screenshot it`s "Microsoft Windows")
- `$caption_XX` sets the title of the window
- `$message_XX` sets the message body of the window

# Kudos
This script is based on https://github.com/Dviros/CredsLeaker.

# Legal
This code is provided for educational use only. If you engage in any illegal activity the author does not take any responsibility for it. By using this code, you agree with these terms.