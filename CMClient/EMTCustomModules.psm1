#===========================================================================
# Functions
#===========================================================================

# Function to run SCCM Client Actions
Function Run-SCCMClientAction{
    [CmdletBinding()]
            
    # Parameters used in this function
    param
    ( 
        [Parameter(Position=0, Mandatory = $True, HelpMessage="Provide server names", ValueFromPipeline = $true)] 
        #[string[]]$Computername,

       [ValidateSet('MachinePolicy', 
                    'DiscoveryData', 
                    'ComplianceEvaluation', 
                    'AppDeployment',  
                    'HardwareInventory', 
                    'UpdateDeployment', 
                    'UpdateScan', 
                    'SoftwareInventory')] 
        [string[]]$ClientAction

    ) 
    $ActionResults = @()
    Try { 
            $ActionResults = Invoke-Command {param($ClientAction)

                    Foreach ($Item in $ClientAction) {
                        $Object = @{} | select "Action name",Status
                        Try{
                            $ScheduleIDMappings = @{ 
                                'MachinePolicy'        = '{00000000-0000-0000-0000-000000000021}'; 
                                'DiscoveryData'        = '{00000000-0000-0000-0000-000000000003}'; 
                                'ComplianceEvaluation' = '{00000000-0000-0000-0000-000000000071}'; 
                                'AppDeployment'        = '{00000000-0000-0000-0000-000000000121}'; 
                                'HardwareInventory'    = '{00000000-0000-0000-0000-000000000001}'; 
                                'UpdateDeployment'     = '{00000000-0000-0000-0000-000000000108}'; 
                                'UpdateScan'           = '{00000000-0000-0000-0000-000000000113}'; 
                                'SoftwareInventory'    = '{00000000-0000-0000-0000-000000000002}'; 
                            }
                            $ScheduleID = $ScheduleIDMappings[$item]
                            Write-Verbose "Processing $Item - $ScheduleID"
                            [void]([wmiclass] "root\ccm:SMS_Client").TriggerSchedule($ScheduleID);
                            $Status = "Success"
                            Write-Verbose "Operation status - $status"
                        }
                        Catch{
                            $Status = "Failed"
                            Write-Verbose "Operation status - $status"
                        }
                        $Object."Action name" = $item
                        $Object.Status = $Status
                        $Object
                    }

        } -ArgumentList $ClientAction -ErrorAction Stop | Select-Object @{n='ServerName';e={$_.pscomputername}},"Action name",Status
    }  
    Catch{
        Write-Error $_.Exception.Message 
    }   
    Return $ActionResults       
    Update-InfoBox -Status $ActionResults    
}

# Function to uninstall the existing CMClient. If trying to run on a remote computer, you must use Copy-CMClientFiles and Start-PSRemoting functions first.
Function Uninstall-CMClient{
    # Tests if CMClient is currently installed. If so, proceeds with the uninstall script.
    if(Test-Path -Path "C:\Windows\CCM","C:\Windows\ccmcache"){
        #Update-InfoBox -Status "Uninstalling CMClient..."

        Write-Host "Beginning CMClient Uninstall..."

        #Stop the CCMExec Service    
        Write-Host "Stopping the CcmExec service..."
        Stop-Service -Name "CcmExec" -Force -ErrorAction SilentlyContinue

        #Run the CCMSetup Uninstall command
        Write-Host "Running the CCMSetup Uninstall command..."
        Start-Process "C:\Temp\CMClient\ccmsetup.exe" -ArgumentList "/uninstall" -Wait

        #Run the CCMClean Tool
        Write-Host "Running the CCMClean Tool..."
        Start-Process "C:\temp\CMClient\ccmclean.exe" -ArgumentList "/all /q" -Wait

        #Stop Services
        $SvcArr = @("BITS","CCMSetup","wuauserv","iphlpsvc","wscsvc","ms enc","Winmgmt")

        ForEach($Svc in $SvcArr){
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            Write-Host "Stopping Services..."
        }

        #Deleting files
        $FoldersArr = @("C:\Windows\ccm","C:\Windows\ccmsetup","C:\Windows\ccmcache","C:\Windows\ccmtemp","C:\SMS","C:\SMSTSLog",
                        "C:\_SMSTaskSequence","C:\Windows\Temp","C:\Windows\sms*.ini","C:\Windows\sms*.mif")

        Foreach($item in $FoldersArr){
            if(Test-Path $item){
                Remove-Item -Path $item -Recurse -Force -Verbose -ErrorAction SilentlyContinue 
                Write-Host "Removing folders..."                   
            }
        }

        #Deleting Reg Keys
        $RegKeysArr = @("HKLM:\Software\Microsoft\ccm","HKLM:\Software\Microsoft\CCMSETUP","HKLM:\Software\Microsoft\SMS",
                        "HKLM:\SYSTEM\CurrentControlSet\Services\CCMSetup","HKLM:\SYSTEM\CurrentControlSet\Services\CcmExec",
                        "HKLM:\SYSTEM\CurrentControlSet\Services\smstsmgr","HKLM:\SYSTEM\CurrentControlSet\Services\CmRcService")

        Foreach($regK in $RegKeysArr){
            if(Test-Path -Path $regK){
                Remove-Item -Path $regK -Recurse -Force -Verbose -ErrorAction SilentlyContinue
                Write-Host "Deleting CMClient Reg Keys..."                 
            }
        }

        #Remove WMI Namespaces
        Write-Host "Removing [CCM] WMI Namespaces..."
        #CCM
        Get-WmiObject -Namespace root -class __Namespace -Filter "name = 'ccm'" | Remove-WMIObject
        Get-WMIObject -namespace "root" -query "SELECT * FROM __Namespace where name = 'ccm'" | remove-wmiobject
        #SMS
        Write-Host "Removing [SMS] WMI Namespaces..."
        Get-WmiObject -Namespace root\cimv2 -class __Namespace -Filter "name = 'sms'" | Remove-WMIObject
        Get-WMIObject -namespace "root" -query "SELECT * FROM __Namespace where name = 'sms'" | remove-wmiobject

        #Remove WMI Objects
        Write-Host "Removing WMI Objects..."
        (Get-WmiObject Win32_Service -filter "name='ccmexec'").Delete()
        (Get-WmiObject Win32_Service -filter "name='ccmsetup'").Delete()
        (Get-WmiObject Win32_Service -filter "name='smstsmgr'").Delete()

        #Reset WMI Repository
        Write-Host "Resetting WMI Repository..."
        Stop-Service -Name "Winmgmt" -Force -ErrorAction SilentlyContinue
        winmgmt /resetrepository

        #Clean Up Winmgmt
        Write-Host "Cleaning up Winmgmt..."
        Remove-Item -Path "C:\Windows\System32\Wbem\Repository.old" -Recurse -Force -ErrorAction SilentlyContinue
        Rename-Item -Path "C:\Windows\System32\Wbem\Repository" -NewName "C:\Windows\System32\Wbem\Repository.old" -Force -ErrorAction SilentlyContinue

        #Clean Up WSUS
        Write-Host "Cleaning up WSUS..."
        $Registrypol= (Test-Path C:\Windows\System32\GroupPolicy\Machine\Registry.pol)
        $RegistrypolOLD= (Test-Path C:\Windows\System32\GroupPolicy\Machine\Registry.pol.OLD)
        $commentcmtx=(Test-Path C:\Windows\System32\GroupPolicy\Machine\comment.cmtx)
        $commentcmtxOLD = (Test-Path C:\Windows\System32\GroupPolicy\Machine\comment.cmtx.OLD)
        $SoftwareDistributionOLD = (Test-Path C:\Windows\SOFTWAREDISTRIBUTION.OLD)
        Get-Service -Name wuauserv | Stop-Service -Force

        IF($Registrypol){
            IF(!($RegistrypolOLD)){
                Rename-item -path "C:\Windows\System32\GroupPolicy\Machine\Registry.pol" -newname "Registry.pol.OLD"
                Write-Host "Renaming [Registry.pol] to [Registry.pol.OLD]"
            }
        }

        IF($commentcmtx){
            IF(!($commentcmtxOLD)){
                Rename-item -path "C:\Windows\System32\GroupPolicy\Machine\comment.cmtx" -newname "comment.cmtx.OLD"
                Write-Host "Renaming [comment.cmtx] to [comment.cmtx.OLD]"
            }
        }

        IF(!($SoftwareDistributionOLD)){
            Remove-Item -Path "C:\Windows\SoftwareDistribution" -Force -Recurse -ErrorAction SilentlyContinue
            Write-Host "Removing [SoftwareDistribution] folder..."
        }

        #Reset MOF Files
        mofcomp 'C:\Program Files\Microsoft Policy Platform\ExtendedStatus.mof'
        Write-Host "Resetting MOF files..."

        Start-Sleep -Seconds 5

        # Run WMIFix Batch File
        Write-Host "Running the WMI Fix Batch File..."
        & "C:\temp\CMClient\WMIFix01.bat"

        #Done
        Write-Host "CMClient Uninstall Complete!"

       # Update-InfoBox -Status "MECM Client has been uninstalled..."
        
    }else{ # if the CMClient is installed, nothing needs to happen.
        #Done
        Write-Host "CMClient is not installed"

        #Update-InfoBox -Status "MECM Client is not currently installed..."

    }


}

# Function to install the CMClient from C:\Temp\CMClient. If trying to run on a remote computer, you must use Copy-CMClientFiles and Start-PSRemoting functions first. 
Function Install-CMClient{
    # Authors: Rachel Catches-Ford and Brandon Kessler
    # Date: 2021.10.26
    # Purpose: Installs the SCCM Client and runs SCCM Client Actions until at least 2 folders are detected in C:\Windows\CCMCache.
    # Updated 2023.05.19 by Rachel Catches-Ford for ACG

    if((Test-Path -path "C:\windows\ccmcache") -or (Test-Path -path "C:\windows\CCM") -eq $true){
        # Tests to see if CMClient is already installed. If so, no further action needed.
        Write-Output "CMClient is already installed..."
        #Update-InfoBox -Status "MECM Client is already installed..."

    }else{
        Write-Output "Installing the CMClient..."
        #Update-InfoBox -Status "Installing the CMClient..."

        # Install SCCM Client
        Write-Output "Installing ConfigMgr Client..."
        $CMGArguments = "/nocrlcheck /mp:HTTPS://ARAPAHOECMG.CO.ARAPAHOE.CO.US/CCM_Proxy_MutualAuth/72057594037927975 CCMHTTPSSTATE=31 CCMHOSTNAME=ARAPAHOECMG.CO.ARAPAHOE.CO.US/CCM_Proxy_MutualAuth/72057594037927975 SMSSiteCode=ACM AADTENANTID=57D7B626-D71D-47F6-84C1-C43BDA19BA16 AADCLIENTAPPID=bf5b6fc2-8019-400c-8c53-add445b4b1a8 AADRESOURCEURI=api://57d7b626-d71d-47f6-84c1-c43bda19ba16/fbac68be-3dda-48df-86b1-89476098667d"
        $arguments = "/mp:ADMMEMCMP1.co.arapahoe.co.us /logon SMSSITECODE=ACM SMSCACHESIZE=20000 /UsePKICert /NoCRLCheck SMSCACHESIZE=20000 /forceinstall"
        Start-Process -FilePath "C:\Temp\CMClient\ccmsetup.exe" -ArgumentList $arguments -Wait
        Start-Process -FilePath "C:\Temp\CMClient\ccmsetup.exe" -ArgumentList $CMGArguments -Wait
    
        $Count = 1
        $ccmcache = 'C:\windows\ccmcache'
        Do{
            Run-SCCMClientAction -ClientAction AppDeployment
            Run-SCCMClientAction -ClientAction MachinePolicy
            Run-SCCMClientAction -ClientAction DiscoveryData
            Run-SCCMClientAction -ClientAction ComplianceEvaluation
            Run-SCCMClientAction -ClientAction HardwareInventory
            Run-SCCMClientAction -ClientAction UpdateDeployment
            Run-SCCMClientAction -ClientAction UpdateScan
            Run-SCCMClientAction -ClientAction SoftwareInventory
            Write-Output "SCCM actions have run $Count times."
            Start-Sleep -Seconds 15
            $Count ++
    
        }While((Get-ChildItem -Path "$ccmcache").count -lt 2)
    
    }
}

# Function to run
Function MECM-RunActions{
    $Count = 1
    
    Do{
        Run-SCCMClientAction -ClientAction AppDeployment
        Run-SCCMClientAction -ClientAction MachinePolicy
        Run-SCCMClientAction -ClientAction DiscoveryData
        Run-SCCMClientAction -ClientAction ComplianceEvaluation
        Run-SCCMClientAction -ClientAction HardwareInventory
        Run-SCCMClientAction -ClientAction UpdateDeployment
        Run-SCCMClientAction -ClientAction UpdateScan
        Run-SCCMClientAction -ClientAction SoftwareInventory
        Write-Output "SCCM actions have run $Count times."
        Start-Sleep -Seconds 15
        $Count ++

    }While($Count -lt 5)
}

# Function to clear the CCMCache folder on the client
Function Clear-CMCache{
    ## Initialize the CCM resource manager com object
    [__comobject]$CCMComObject = New-Object -ComObject 'UIResource.UIResourceMgr'

    ## Get the CacheElementIDs to delete
    $CacheInfo = $CCMComObject.GetCacheInfo().GetCacheElements()

    ## Remove cache items
    ForEach ($CacheItem in $CacheInfo) {
        $null = $CCMComObject.GetCacheInfo().DeleteCacheElement([string]$($CacheItem.CacheElementID))
    }
}

Export-ModuleMember -Function Run-SCCMClientAction
Export-ModuleMember -Function Uninstall-CMClient
Export-ModuleMember -Function Install-CMClient
Export-ModuleMember -Function MECM-RunActions
Export-ModuleMember -Function Clear-CMCache