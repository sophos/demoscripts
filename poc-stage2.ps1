schtasks /create /sc onlogon /tn "POC Stage 3" /tr "c:\threat\poc-stage3.bat" /ru Sophos /F
Start-Process c:\threat\SophosSetup.exe --quiet -Wait
schtasks /delete /tn "POC Stage 2" /F
schtasks /create /sc minute /mo 1 /tn "POC EAP Watcher" /tr "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe -ExecutionPolicy Bypass c:\threat\poc-eap_watcher.ps1" /ru System /F