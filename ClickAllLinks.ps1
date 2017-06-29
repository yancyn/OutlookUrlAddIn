Clear-Host
# TODO: idle time 1000ms = 1s
$idle = 1000

# Parse from urls.txt
# $urls = @("http://news.google.com", "https://www.yahoo.com/news/")
$urls = Get-Content url.txt

$ie = New-Object -COMObject InternetExplorer.Application
$ie.Visible = $True
Add-Type -AssemblyName System.Windows.Forms

# Start loop collection of url list and navigate
Add-Type -AssemblyName PresentationFramework
# For($i=0;$i -lt $urls.Length;$i++) {
foreach($url in $urls) {
    Start-Sleep -s 1
    $ie.Navigate($url)
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}");Start-Sleep -m $idle
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");Start-Sleep -m $idle
}
$ie.Quit()
[System.Windows.MessageBox]::Show("Done", "Auto Submit")
