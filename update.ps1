# Windows 11 SE Resources Image Update Script
# v0.3

$ErrorActionPreference = "Stop"

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
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion, latest version is $latestVersion. No updates needed." | Out-file $logFile -append
} elseif ($currentVersion -lt $latestVersion) {
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion, latest version is $latestVersion. Starting updates." | Out-file $logFile -append

# UPDATE COMMANDS GO HERE

Try {
    Write-Output "$(Get-TimeStamp) Download update zip file" | Out-file $logFile -append
    Invoke-WebRequest -Uri "http://threatmenu.naseeast.com/threat/update/update.zip" -OutFile "c:\threat\image\update\update.zip" 2>&1 | Tee-Object -FilePath "$logFile" -Append
    Write-Output "$(Get-TimeStamp) File downloaded successfully." | Out-File -FilePath $logFile -Append

    Expand-Archive -Path "c:\threat\image\update\update.zip" -DestinationPath "c:\threat\image\update\" -Force
    Write-Output "$(Get-TimeStamp) File unzipped successfully." | Out-File -FilePath $logFile -Append

    Copy-Item -Path "C:\threat\image\update\Feature Control Script.bat" -Destination "C:\threat\image"
    Copy-Item -Path "C:\threat\image\update\Feature Control.lnk" -Destination "C:\users\sophos\Desktop"
    
}
Catch {
    $errorMessage = $_.Exception.Message
    Write-Output "$(Get-TimeStamp) ERROR: $errorMessage" | Out-File -FilePath $logFile -Append
    exit 1
}

} else {
	Write-Output "$(Get-TimeStamp) ERROR Version not found in Registry" | Out-file $logFile -append
}  
