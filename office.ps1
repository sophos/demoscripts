#Name our event source
$EventSource = “OfficeLStatus”
 
#Create an event source in event viewer to log any failures.
New-EventLog –LogName Application –Source $EventSource -ErrorAction SilentlyContinue
 
#Function to clean up our event log provider so we don't leave it behind when we're done
function Remove-OfficeLicenseActivationEventProvider()
{
    #Remove our eventlog provider
    Remove-EventLog -Source $EventSource
}
 
#Event for displaying license status
function Write-LicenseStatusEventLog($ID,$Type,$Message)
{
    Write-EventLog –LogName Application –Source “OfficeLStatus” –EntryType $Type –EventID $ID –Message “$Message” -ErrorAction SilentlyContinue
}
 
#Function to find the correct version of OSPP to use
function Get-OfficeSoftwareProtectionPlatform
{
     
    [CmdletBinding()][OutputType([string])]
    param ()
       
    $File = Get-ChildItem $env:ProgramFiles"\Microsoft Office" -Filter "OSPP.VBS" -Recurse -ErrorAction SilentlyContinue
    if (($File -eq $null) -or ($File -eq ''))
    {
        $File = Get-ChildItem ${env:ProgramFiles(x86)}"\Microsoft Office" -Filter "OSPP.VBS" -Recurse -ErrorAction SilentlyContinue
        $File = $File.FullName
    }
 
    if (($File -eq $null) -or ($File -eq ''))
    {
        Write-LicenseStatusEventLog -ID "9010" -Type "ERROR" -Message "Unable to find Office software protection platform!"
        $File = $false
    }
 
    return $File
}
 
#Function to find the SoftwareLicenseManager
function Get-SoftwareLicenseManager
{
    $File = Get-ChildItem $env:windir"\system32" | Where-Object { $_.Name -eq "slmgr.vbs" }
    $File = $File.FullName
   
    if(!$File)
    {
        Write-LicenseStatusEventLog -ID "9010" -Type "ERROR" -Message "Failed to find the Software License Manager script!!!`nResolve this before attempting the script again!"
        $File = $null
    }
 
    return $File
}
 
#Function to set the office product key in OSPP
function Set-OfficeProductKey($OSPP)
{
    $Errors = $false
   
    #Specify an office product key, this can be done within the script or in a text file and use Get-Content to read said text file
    #For testing I am using a single static key
    $OfficeProductKey = "88T6N-FFP6T-9PCWW-QG66F-YTDVC"
    Write-LicenseStatusEventLog -ID 9003 -Type "INFORMATION" -Message "Setting office product key to: $OfficeProductKey"
   
    $Executable = $env:windir + "\System32\cscript.exe"
    $Switches = [char]34 + $OSPP + [char]34 + [char]32 + "/inpkey:" + $OfficeProductKey
 
    if ((Test-Path $Executable) -eq $true)
    {
        $ErrCode = (Start-Process -FilePath $Executable -ArgumentList $Switches -Wait -Passthru).ExitCode
    }
 
    if (($ErrCode -eq 0) -or ($ErrCode -eq 3010))
    {
        Write-LicenseStatusEventLog -ID 9003 -Type "INFORMATION" -Message "Successfully set productkey within OSPP"
    }
    else
    {
        Write-LicenseStatusEventLog -ID 9003 -Type "ERROR" -Message "Failed to set productkey within OSPP!!!`nErrorCode:$ErrCode"
        $Errors = $true
    }
}
 
#Function to activate office
function Invoke-OfficeActivation($OSPP)
{
    
    $Errors = $false
   
    Write-LicenseStatusEventLog -ID "9004" -Type "INFORMATION" -Message "Attempting to activate office."
 
   
    $Executable = $env:windir + "\System32\cscript.exe"
    #$Switches = "`"$OSPP`" /act"
    #$Switches = $Switches.ToString()
 
   
    if ((Test-Path $Executable) -eq $true)
    {
        #We should do this a different way using the $Executable and $switches inside of start process like below but, that isn't working for me...
        #$ErrCode = (Start-Process -FilePath $Executable -ArgumentList "$Switches" -Wait -WindowStyle Minimized -Passthru).ExitCode
        $OSPPActRet = C:\Windows\System32\cscript.exe "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /act
    }
 
    if($OSPPActRet)
    {
        $OSPPActRet = $OSPPActRet | Select-String "<Product activation successful>"
        $OSPPActRet = $OSPPActRet.ToString()
    }
 
    if ($OSPPActRet -eq "<Product activation successful>")
    {
        Write-LicenseStatusEventLog -ID "9005" -Type "INFORMATION" -Message "Successfully activated office."
    }
    else
    {
        Write-LicenseStatusEventLog -ID "9005" -Type "ERROR" -Message "Failed to activate office!`nErrorCode:$ErrCode"
        $Errors = $true
    }
    return $Errors
}
 
#Log and event for the initalization
Write-LicenseStatusEventLog -ID "9000" -Type "INFORMATION" -Message "Office license activation validation started."
 
#Find the OSPP version
$OSPP = Get-OfficeSoftwareProtectionPlatform
 
#Find the SoftwareLicenseManager
$SLM = Get-SoftwareLicenseManager
 
if(!$OSPP)
{
    #Failed to find a Office Software Protection Platform VBS script, either we don't have access or office isn't installed...
    Write-LicenseStatusEventLog -ID "9911" -Type "ERROR" -Message "Office installation not found!!!`nRemoving script event provider and bailing!"
    Remove-OfficeLicenseActivationEventProvider
    return
}
else
{
    if(!$SLM)
    {
        #Failed to find the  Software license manager, just clean up the event provider and bail
        Remove-OfficeLicenseActivationEventProvider
        return
    }
 
    #Log an event that we found OSPP and the location
    Write-LicenseStatusEventLog -ID "9001" -Type "INFORMATION" -Message "Found OSPP at: $OSPP"
 
    #Now check the validation status
    $cscriptOutput = cscript $OSPP /dstatus
 
    #We don't care about most of the output of the csript... snatch only the lines we care about
    $licenseStatus = $cscriptOutput |Select-String -SimpleMatch "LICENSE STATUS:"
 
    if($licenseStatus.Count -gt 1)
    {
        #We have multiple licenses installed... I don't have the logic to handle that so write an event and bail
        Write-LicenseStatusEventLog -ID "9999" -Type "ERROR" -Message "MULTIPLE OFFICE INSTALLS FOUND!`nScript does not currently have logic to handle this scenario so bailing out!"
        Remove-OfficeLicenseActivationEventProvider
        return
    }
    elseif($licenseStatus -eq $null)
    {
        #No license found!
        Write-LicenseStatusEventLog -ID "9002" -Type "INFORMATION" -Message "No currently installed license found!"
    }
    else
    {
        $licenseStatus = $licenseStatus.ToString()
        $licenseStatus = $licenseStatus.Substring($licenseStatus.IndexOf("  -")).Replace("-","")
        $licenseStatus = $licenseStatus.Replace(" ","")
        #Log the license status to eventlog
        Write-LicenseStatusEventLog -ID "9002" -Type "INFORMATION" -Message "Office license activation status:$licenseStatus"
    }
 
   
    #If we're anything other than "LICENSED" let's go ahead and activate
    if($licenseStatus -ne "LICENSED")
    {
 
        #Attempt to license, for this we are going to need a product key. We can either manually populate that here OR we can pull it from a text file, this is up to you.
        $setKeyStatusFailed = Set-OfficeProductKey($OSPP)
 
        if($setKeyStatusFailed)
        {
            #If we're in here, setting the key failed... Set-OfficeProductKey already threw an event in the event log, clean up and bail
            Remove-OfficeLicenseActivationEventProvider
            return
        }
 
        #We successfully set the key, let's go activate
        $newActivationStatusFailed = Invoke-OfficeActivation
 
        #We're done at this point. Regardless if we failed or succeeded just clean up and bail
        Remove-OfficeLicenseActivationEventProvider
        return
    }
    else
    {
        Write-LicenseStatusEventLog -ID "9020" -Type "INFORMATION" -Message "Office license already activated.`nCleaning up script event provider and bailing`nNo additional work necessary."
        Remove-OfficeLicenseActivationEventProvider
        return
    }
}  