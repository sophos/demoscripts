Function eapcheck {
$keyPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Sophos\Management\Policy\ThreatProtection\*\amsi_protection"
$keyVal = "amsi_protection_block_on_detect"
$check = Get-ItemProperty -Path $keyPath | Select $keyVal
#write-host $newCheck.$keyVal
if($check.$keyVal -eq 0) {
schtasks /delete /tn "POC EAP Watcher" /F
Restart-Computer -force
}
}
eapcheck