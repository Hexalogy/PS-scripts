taskkill /F /IM "chrome.exe"
Start-Sleep -Seconds 5

Remove-Item "C:\Users\$SAMID\AppData\Local\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -EA SilentlyContinue
Remove-Item "C:\Users\$SAMID\AppData\Local\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -EA SilentlyContinue
Remove-Item "C:\Users\$SAMID\AppData\Local\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -EA SilentlyContinue
Remove-Item "C:\Users\$SAMID\AppData\Local\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -EA SilentlyContinue

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Cache has been cleared. Please relaunch Chrome again.", 0, "Chrome", 0)
