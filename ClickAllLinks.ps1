Clear-Host
$pause = 500
# TODO: Parse from urls.txt
$urls = @("http://news.google.com", "https://www.yahoo.com/news/")
$ie = New-Object -COMObject InternetExplorer.Application
$ie.Visible = $True
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
For($i=0;$i -lt $urls.Length;$i++) {
    # [System.Windows.MessageBox]::Show($i)
    Start-Sleep -s 1
    $ie.Navigate($urls[$i])
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}");Start-Sleep -m $pause
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");Start-Sleep -m $pause
}
$ie.Quit()
