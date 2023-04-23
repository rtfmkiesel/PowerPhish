# Variants
Here is a collection of variants that I did some experiments with.

## PowerPhish_Browsers
This script scans the running tasks for Firefox, Chrome and MS Edge. After it finds one running, the script will launch the window with the browser name as the target. 

![PowerPhish Browsers](/Screenshots/variant_browsers.PNG)

## PowerPhish_Outlook
This variant will scan the default Outlook Folder for .pst and .ost files. The script simply extracts the e-mail name from the file name itself. After that, the script waits for a running Outlook process. If Outlook is running, the script will present the windows with the e-mail address as target.

![PowerPhish Outlook](/Screenshots/variant_outlook.PNG)