# Windows 11 SE Resources Image Update Script
# v0.2

# Define latest image version
$latestVersion = "2"

# Specify registry path and value
$registryPath = "HKLM:\SYSTEM\SophosSE"
$registryValue = "Version"

# Function to gete currrent time for logging
function Get-TimeStamp {
    
    return "[{0:yyyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    
}

# Specify log file location
$logFile = "c:\threat\image\update_log.txt"

# Script start
Write-Output "$(Get-TimeStamp) Update script starting" | Out-file $logFile -append

# Get $currrentVersion from Registry key
$currentVersion = Get-ItemPropertyValue -Path $registryPath -Name $registryValue -ErrorAction SilentlyContinue

if ($currentVersion -eq "BETA") {
	Write-Output "$(Get-TimeStamp) BETA CHANNEL, applying changes" | Out-file $logFile -append
# BETA UPDATE COMMANDS GO HERE

}  elseif ($currentVersion -eq $latestVersion) {
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion" | Out-file $logFile -append
	Write-Output "$(Get-TimeStamp) Latest version is $latestVersion, no updates needed" | Out-file $logFile -append
} elseif ($currentVersion -lt $latestVersion) {
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion" | Out-file $logFile -append
	Write-Output "$(Get-TimeStamp) Latest version is $latestVersion, starting updates" | Out-file $logFile -append
# UPDATE COMMANDS GO HERE
Write-Output "$(Get-TimeStamp) Download feature control script" | Out-file $logFile -append
Invoke-WebRequest -Uri "http://threatmenu.naseeast.com/threat/update/assets/Feature Control Script.bat" -OutFile "c:\threat\image\Feature Control Script.bat" > $logFile

Write-Output "$(Get-TimeStamp) Download feature control shortcut" | Out-file $logFile -append
Invoke-WebRequest -Uri "http://threatmenu.naseeast.com/threat/update/assets/Feature Control.lnk" -OutFile "c:\users\sophos\desktop\Feature Control.lnk" > $logFile

} else {
	Write-Output "$(Get-TimeStamp) ERROR Version not found in Registry" | Out-file $logFile -append
} 
