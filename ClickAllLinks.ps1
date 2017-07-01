param([Int32]$idle=1000)
#idle time parameter from command prompt. idle time 1000ms = 1s

# Wrap from url.txt into array.
# Make sure source file url.txt must separate url by new line
# $urls = @("http://news.google.com", "https://www.yahoo.com/news/")
$urls = Get-Content url.txt

$ie = New-Object -COMObject InternetExplorer.Application
$ie.Visible = $True # set to false to run at background
Add-Type -AssemblyName System.Windows.Forms

# Start loop collection of url list and navigate
foreach($url in $urls) { # For($i=0;$i -lt $urls.Length;$i++)
    Start-Sleep -s 1
    $ie.Navigate($url)
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}");Start-Sleep -m $idle
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");Start-Sleep -m $idle
}
$ie.Quit()

Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show("Done", "Auto Submit")
