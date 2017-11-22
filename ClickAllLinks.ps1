param([Int32]$idle=500)
#idle time parameter from command prompt. idle time 1000ms = 1s

# Wrap from url.txt into array.
# Make sure source file url.txt must separate url by new line
# $urls = @("http://news.google.com", "https://www.yahoo.com/news/")
$urls = Get-Content url.txt

$ie = New-Object -COMObject InternetExplorer.Application
$ie.Visible = $True # set to false to run at background

# Example of submit in google search
Function SearchInGoogle {
  $doc = $ie.Document
  $txt = $doc.getElementById("lst-ib")
  $num1 = Get-Random
  $num2 = Get-Random
  $num1Str = "{0}" -f $num1
  $num2Str = "{0}" -f $num2
  $txt.value = -join($num1Str, "+", $num2Str)

  $form = $doc.getElementById("tsf")
  $form.submit()
}

# Start loop collection of url list and navigate
# For($i=0;$i -lt $urls.Length;$i++)
foreach($url in $urls) {
    # see http://www.westerndevs.com/simple-powershell-automation-browser-based-tasks/
    $ie.Navigate($url)
    while($ie.Busy) { Start-Sleep -s 1 }

    $doc = $ie.Document
    SearchInGoogle
    while($doc.readyState -ne "complete") { Start-Sleep -s 1 }
}
# $ie.Quit()

Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show("Done", "Auto Submit")
