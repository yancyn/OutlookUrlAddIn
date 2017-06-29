# Outlook Extract Url Add-In
Extract url from outlook folder.

## Download
Under [release](https://github.com/yancyn/OutlookUrlAddIn/releases) tab pick the latest version.

## How to Install Outlook Add-In
Double click on OutlookUrlAddIn.vsto.

## How to Remove Outlook Add-In (VSTO)
Outlook > File menu > Options > Add-Ins > Go > OutlookUrlAddIn > Remove.

## How To Use Extract Link from Outlook Folder
1. Double click OutlookUrlAddIn.vsto. This will install to your Outlook 2013.
2. Under local folder > select specified folder > Add-Ins.
3. Press on 'Get Url' on (ribbon) menu.
4. A new text file will pop up after finished processing.
5. Done.

## How to Auto Submit Links
1. Download ```ClickAllLinks.ps1``` and ```Start.bat```.
2. Prepare the url list in a file (see step _How To Extract Link_ above) and save as ```url.txt``` same location in step 1.
3. Double click on ```Start.bat``` or manually start the powershell script.
```
cmd > powershell -windowstyle hidden -ExecutionPolicy ByPass -File "ClickAllLinks.ps1"
```
4. Done.

## References
- https://www.codeproject.com/Articles/1112815/How-to-Create-an-Add-in-for-Microsoft-Outlook
- http://regexr.com/38l0t
- https://msdn.microsoft.com/en-us/powershell/scripting/getting-started/cookbooks/creating-.net-and-com-objects--new-object-
- http://www.tomsitpro.com/articles/powershell-for-loop,2-845.html
- https://social.technet.microsoft.com/Forums/Azure/en-US/78d5a5fa-bb82-4c2d-a2c1-96d518b9bd74/need-to-read-text-file-as-an-array-and-get-elements-from-each-line-of-array-using-powershell?forum=winserverpowershell