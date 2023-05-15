# Windows 11 SE Resources Image Update Script
# v0.2

# Define latest image version
$latestVersion = "3"

# Specify registry path and value
$registryPath = "HKLM:\SYSTEM\SophosSE"
$registryValue = "Version"

# Function to gete currrent time for logging
function Get-TimeStamp {
    
    return "[{0:yyyy-MM-dd} {0:HH:mm:ss}]" -f (Get-Date)
    
}

# Script start
Write-Output "$(Get-TimeStamp) Update script starting" | Out-file c:\threat\image\update_log.txt -append

# Get $currrentVersion from Registry key
$currentVersion = Get-ItemPropertyValue -Path $registryPath -Name $registryValue -ErrorAction SilentlyContinue

if ($currentVersion -eq "BETA") {
	Write-Output "$(Get-TimeStamp) BETA CHANNEL, applying changes" | Out-file c:\threat\image\update_log.txt -append
# BETA UPDATE COMMANDS GO HERE

}  elseif ($currentVersion -eq $latestVersion) {
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion" | Out-file c:\threat\image\update_log.txt -append
	Write-Output "$(Get-TimeStamp) Latest version is $latestVersion, no updates needed" | Out-file c:\threat\image\update_log.txt -append
} elseif ($currentVersion -lt $latestVersion) {
	Write-Output "$(Get-TimeStamp) Image version is $currentVersion" | Out-file c:\threat\image\update_log.txt -append
	Write-Output "$(Get-TimeStamp) Latest version is $latestVersion, starting updates" | Out-file c:\threat\image\update_log.txt -append
# UPDATE COMMANDS GO HERE

} else {
	Write-Output "$(Get-TimeStamp) ERROR Version not found in Registry" | Out-file c:\threat\image\update_log.txt -append
} 
