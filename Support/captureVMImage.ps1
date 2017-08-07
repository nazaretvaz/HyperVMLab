[CmdletBinding()]
param (
    [Parameter(Mandatory, Position=0)]  
    [String] $VMName,

    [Parameter(Mandatory, Position=1)]
    [pscredential] $VMCredential
)

#Set verbose preference; This is needed for verbose output from remote session
$VerbosePreference = 'Continue'

#UnattendXML template
$unattendXml = @'
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend"
        xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <settings pass="generalize">
        <component name="Microsoft-Windows-PnpSysprep"
                processorArchitecture="amd64"
                publicKeyToken="31bf3856ad364e35"
                language="neutral" versionScope="nonSxS">
            <PersistAllDeviceInstalls>false</PersistAllDeviceInstalls>
        </component>
    </settings>
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-International-Core"
                processorArchitecture="amd64"
                publicKeyToken="31bf3856ad364e35"
                language="neutral" versionScope="nonSxS">
            <InputLocale>{0}</InputLocale>
            <SystemLocale>{0}</SystemLocale>
            <UILanguage>{0}</UILanguage>
            <UserLocale>{0}</UserLocale>
        </component>
        <component name="Microsoft-Windows-Shell-Setup"
                processorArchitecture="amd64"
                publicKeyToken="31bf3856ad364e35"
                language="neutral" versionScope="nonSxS">
                <UserAccounts>
                   <AdministratorPassword>
                      <Value>{1}</Value>
                      <PlainText>true</PlainText>
                   </AdministratorPassword>
                </UserAccounts>
                <OOBE>
                   <HideEULAPage>true</HideEULAPage>
                </OOBE>
        </component>
    </settings>
</unattend>
'@

#Guest KVP path
$guestKVP = 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters'

if ((Get-ItemProperty -Path $guestKVP -ErrorAction SilentlyContinue).VirtualMachineName -eq $VMName)
{
    Write-Verbose -Message 'Script is running inside the guest ...'
    
    #Start executing the capture image process
    Write-Verbose -Message 'Preparing and copying unattend.xml ...'
    $systemLocale = (Get-UICulture).Name
    $adminstratorPassword = $VMCredential.GetNetworkCredential().Password
    ($unattendXml -f $systemLocale, $adminstratorPassword) | Out-File C:\unattend.xml -Encoding utf8 -Force

    #Start Sysprep
    $sysprepCommandline = '/generalize /oobe /mode:vm /unattend:c:\unattend.xml /quiet /quit'
    Write-Verbose -Message 'Starting sysprep ...'
    Start-Process -FilePath "${env:SystemRoot}\System32\Sysprep\Sysprep.exe" -ArgumentList $sysprepCommandline -Wait -Verbose
}
else
{
    #We are on the host. We need to bootstrap this script.
    Write-Verbose -Message 'Script is running on the host ...'
    
    #Check if Hyper-V PowerShell module is available or not
    if (-not (Get-Module -ListAvailable -Name Hyper-V))
    {
        throw 'Hyper-V module is not available.'
    }

    #Check if the VM is running on the local system
    $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
    if (-not ($vm))
    {
        throw "${VMName} not found on the host."
    }
    else
    {
        #Check if the VM is in running state or not; if not start VM.
        if ($vm.State -ne 'Running')
        {
            if ($vm.State -eq 'Off')
            {
                Write-Verbose -Message 'Starting virtual machine ...'
                Start-VM -VM $vm
            }
            elseif ($vm.State -eq 'Paused')
            {
                Write-Verbose -Message 'Resuming virtual machine ...'
                Resume-VM -VM $vm
            }

            #Wait for a few seconds to ensure the IC comes up
            Write-Verbose -Message 'Waiting for VM integration components ...'
            do
            {
                Start-Sleep -milliseconds 100
            } until ((Get-VMIntegrationService -VM $vm | ?{$_.name -eq "Heartbeat"}).PrimaryStatusDescription -eq "OK")
        }

        #Check if Guest Service interface is working or not; if not enable
        if ((Get-VMIntegrationService -VM $vm -Name 'Guest Service Interface' -Verbose).Enabled -eq $false)
        {
            #Enable GSI
            Write-Verbose -Message 'Guest Service Interface is not enabled. Enabling this.'
            Enable-VMIntegrationService -Name 'Guest Service Interface' -VM $vm -Verbose

            #Wait for a bit for the service to start
            Start-Sleep -Seconds 5
        }

        #Copy the script
        try
        {
            $scriptPath = $MyInvocation.MyCommand.Path
            $scriptName = $MyInvocation.MyCommand.Name
            Write-Verbose -Message 'Copying the script and bootstraping capture process.'
            Copy-VMFile -VM $vm -SourcePath $scriptPath -DestinationPath "C:\Scripts\${scriptName}" -FileSource Host -CreateFullPath -Force -Verbose -ErrorAction Stop
        }
        catch
        {
            Write-Error $_
        }

        #Invoke script
        try
        {
            Write-Verbose -Message 'Invoking the script inside the guest!'
            $vmSession = New-PSSession -VMName $VMName -Credential $VMCredential        
            Invoke-Command -Session $vmSession -FilePath "C:\Scripts\${scriptName}" -ArgumentList $VMName, $VMCredential -Verbose

            #if Sysprep was successfully, we need to check the exitcode from the process
            $oobeStatus = Invoke-Command -Session $vmSession -ScriptBlock { (Get-ItemProperty -Path HKLM:\SYSTEM\Setup -Name OOBEInProgress).OOBEInProgress }

            if ($oobeStatus -eq 1)
            {
                Write-Verbose -Message 'Sysprep completed and VM will shutdown now.'
                Remove-PSSession -Session $vmSession
                Stop-VM -Name $VMName -Force
            }
            else
            {
                Write-Error -Message 'Sysprep seemed to have failed. Check in the VM'
            }
        }
        catch
        {
            Write-Error $_
        }
    }
}
