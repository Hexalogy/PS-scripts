#write-host "If theres an error, make sure to check if the password is updated`n" -ForegroundColor Yellow

$Credential = Get-Credential
$UserName = $Credential.UserName
$Password = $Credential.GetNetworkCredential().Password

$hname = Read-Host "Hostname"

write-host "Getting current processes on target machine.."

#call pslist for processes
\\sltcp-fps1\software\powershell\ees\pslist.exe -t -s 1 \\$hname

Write-Host "Computer name is $hname `n"
  "-------------------------------------------------------`n"

$pname = Read-Host "Process Name to kill (exclude .exe)"

\\sltcp-fps1\software\powershell\ees\pskill.exe -t \\$hname $pname -u $UserName -p $Password

Read-Host -Prompt "Press Enter to exit"
