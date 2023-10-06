# Command Center Desktop
# Created by Josh Spring, modified by Rachel Catches-Ford
# Last modified 2023.07.05
# Version 3

$inputXML = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ACG Command Center 2.0" Height="678" Width="759" Icon="C:\ProgramData\ACG\CommandCenter\ACGicon.ico" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen" FontFamily="Tahoma">
    <Window.Background>
        <ImageBrush ImageSource="C:\ProgramData\ACG\CommandCenter\EyeRelaxer.jpg"/>
    </Window.Background>
    <Grid Height="NaN" Margin="0,0,0,-6">
        <TextBox x:Name="ComputerNameBox" HorizontalAlignment="Center" Margin="0,41,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="166" Height="33" FontSize="18" Foreground="#FF008FFD" FontWeight="Bold" TextAlignment="Center">
            <TextBox.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </TextBox.Background>
        </TextBox>
        <Button x:Name="C_DriveButton" Content="C$" HorizontalAlignment="Left" Margin="22,116,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="InstallMECMClientButton" Content="Install MECM Client" HorizontalAlignment="Left" Margin="185,384,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="Green" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="UserButton" Content="Query User" HorizontalAlignment="Left" Margin="170,116,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="RDPbutton" Content="Remote Desktop" HorizontalAlignment="Left" Margin="170,209,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold" RenderTransformOrigin="-0.652,-1.006">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="MECM_Console_Button" Content="MECM Console" HorizontalAlignment="Left" Margin="435,116,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="GPO_Button" Content="Group Policy" HorizontalAlignment="Left" Margin="585,116,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="ClearButton" Content="Clear" HorizontalAlignment="Left" Margin="462,46,0,0" VerticalAlignment="Top" Width="60" Height="23" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Normal" RenderTransformOrigin="0.142,0.578" FontFamily="Tahoma">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Label x:Name="ComputerLabel" Content="Computer Name" HorizontalAlignment="Center" Margin="0,2,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.38,-0.32" Width="174" Foreground="#FFFF0606" FontSize="20" FontWeight="Bold" FontFamily="Tahoma" Background="Black"/>
        <Button x:Name="Powershell_LA_Button" Content="Powershell (LA)" HorizontalAlignment="Left" Margin="435,210,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Label x:Name="InfoLabel1" Content="Remote Tools" HorizontalAlignment="Left" Margin="88,67,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.38,-0.32" Width="152" Foreground="#FFFF0606" FontSize="20" FontWeight="Bold" Background="Black"/>
        <Label x:Name="InfoLabel2" Content="Admin Tools" HorizontalAlignment="Left" Margin="527,67,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.38,-0.32" Width="134" Foreground="#FFFF0606" FontSize="20" FontWeight="Bold" Background="Black"/>
        <TextBox x:Name="InfoBox1" HorizontalAlignment="Center" Margin="0,478,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="150" Foreground="#FF1A9CEF">
            <TextBox.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </TextBox.Background>
        </TextBox>
        <Button x:Name="PingButton" Content="Ping" HorizontalAlignment="Left" Margin="22,163,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="PowershellButton" Content="Powershell" HorizontalAlignment="Left" Margin="435,256,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="PrintButton" Content="Print Mgmt" HorizontalAlignment="Left" Margin="435,163,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="CompMgmtButton" Content="Comp Mgmt" HorizontalAlignment="Left" Margin="22,209,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="VSCodeButton" Content="VS Code" HorizontalAlignment="Left" Margin="584,256,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="TerminalButton" Content="Terminal" HorizontalAlignment="Left" Margin="435,303,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="HyperVButton" Content="Hyper-V" HorizontalAlignment="Left" Margin="584,210,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="DiskMgmtButton" Content="Disk Mgmt" HorizontalAlignment="Left" Margin="584,163,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="GitHubButton" Content="GitHub Desktop" HorizontalAlignment="Left" Margin="585,303,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="RemoveFromMECMButton" Content="Remove From MECM" HorizontalAlignment="Left" Margin="106,431,0,0" VerticalAlignment="Top" Width="154" Height="42" Foreground="#FFFF0606" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="RemoveFromADButton" Content="Remove From AD" HorizontalAlignment="Left" Margin="509,431,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FFFF0606" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="UninstallMECMClientButton" Content="Uninstall MECM Client" HorizontalAlignment="Left" Margin="14,384,0,0" VerticalAlignment="Top" Width="166" Height="42" Foreground="#FFFF0606" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="QueryADOUButton" Content="Query AD OU" HorizontalAlignment="Left" Margin="170,163,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="MECMAllActionsButton" Content="Run ALL Actions" HorizontalAlignment="Left" Margin="14,342,0,0" VerticalAlignment="Top" Width="166" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Button x:Name="ClearMECMCacheButton" Content="Clear Client Cache" HorizontalAlignment="Left" Margin="185,342,0,0" VerticalAlignment="Top" Width="150" Height="42" Foreground="#FF1A9CEF" FontSize="14" FontWeight="Bold">
            <Button.Background>
                <SolidColorBrush Color="Black" Opacity="0.8"/>
            </Button.Background>
        </Button>
        <Label x:Name="InfoLabel3" Content="MECM Tools" HorizontalAlignment="Left" Margin="116,299,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.38,-0.32" Width="134" Foreground="#FFFF0606" FontSize="20" FontWeight="Bold" Background="Black"/>

    </Grid>
</Window>

"@ 
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$xaml = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | Foreach-Object {"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop
    }catch{
        throw
    }
}
 
Function Get-FormVariables{
    if ($global:ReadmeDisplay -ne $true){
        Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true
    }

    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}
 
# Use Get-FormVariables to see the list of variables
 
#===========================================================================
# Add code to the various form elements in GUI
#===========================================================================

# Identify logged on user and add LA to end of username
$ExplorerProcess = Get-WmiObject win32_process | Where-Object { $_.name -Match 'explorer'}

$LoggedOnUser = if($ExplorerProcess.getowner().user.count -gt 1){
    $ExplorerProcess.getowner().user[0]
    }else{
        $ExplorerProcess.getowner().user
    }
#!!!!!!!!# Update your user account - this simply takes the current user account and adds 'LA' to the end of it. Modify for your needs
$UserLAaccount = $LoggedOnUser + "LA"

# Prevent Command Center from continuing unless an "IT" user account is logged in
If ($LoggedOnUser -notlike "it*"){
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.Forms.MessageBox]::Show('Only IT department users can run this program')
    Exit 1
}

#!!!!!!!!# Collect credentials - Update your Domain and your Admin account if you want it to pre-populate 
#!!!!!!!!# $LACreds is used throughout this script to run programs as this Admin Account
$LAcreds = Get-Credential Domain\$UserLAaccount -Message "Enter your Admin account username/password"

# Load MECM module
Set-Location 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin'
Import-Module .\ConfigurationManager.psd1

#===========================================================================
# Functions
#===========================================================================

# Function to update the Command Center information box
Function Update-InfoBox{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Status to display in the Info Box.")]
        [string]$Status
    )

    # Clear info box
    $WPFInfoBox1.Clear()

    # Display info in the info box
    $Output1 = [pscustomobject]$Status
    $WPFInfoBox1.AddText("$Output1")
}

# Function: Use Query User to verify computer is actually online and the DNS record is accurate
Function Verify-Computer {
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear variables
    Clear-Variable -Name Query -ErrorAction SilentlyContinue
    Clear-Variable -Name Status -ErrorAction SilentlyContinue
    Clear-Variable -Name LASTEXITCODE -ErrorAction SilentlyContinue
    Clear-Variable -Name Proceed -ErrorAction SilentlyContinue
    
    # Use quser to test if computer is actually online. Helps verify DNS record is accurate as well.

    $Query = quser /server:$Computer 2>&1

            Switch -Wildcard ($Query) {
            '*[5]:Access is denied*' {

                # Info to display in infobox
                $Status = 'DNS problem' # Access denied. DNS problem. Exit code 1.

                Set-Variable -Name Proceed -Value $False -Scope Global
                Write-Host "Proceed is set to $Proceed"

                Write-Host "DNS problem found when attemping to connect to $Computer"
                }

            '*No User exists for*' {

                # Info to display in infobox
                $Status = 'Connecting to computer (with no user logged on currently)' # No user logged on. Exit code 1

                Set-Variable -Name Proceed -Value $True -Scope Global
                Write-Host "Proceed is set to $Proceed"

                Write-Host "Attempting to connect to $Computer with no logged on user"
                }

            '*console*' {

                # Info to display in infobox
                $Status = 'Connecting to computer (with a user currently logged on locally)' # User logged on. Exit code 1.

                Set-Variable -Name Proceed -Value $True -Scope Global
                Write-Host "Proceed is set to $Proceed"
                
                Write-Host "Attempting to connect to $Computer with logged on user"
                }

            '*rdp*' {

                # Info to display in infobox
                $Status = 'Connecting to computer (with a user currently logged on through RDP)' # User logged on. Exit code 1.

                Set-Variable -Name Proceed -Value $True -Scope Global
                Write-Host "Proceed is set to $Proceed"
                
                Write-Host "Attempting to connect to $Computer with logged on RDP user"
                }

            default {

                # Info to display in infobox
                $Status = 'Unable to connect to computer. Possible reasons: Computer offline, DNS problems, misspelled name'
                
                Set-Variable -Name Proceed -Value $False -Scope Global
                Write-Host "Proceed is set to $Proceed"
            }
          }
}

# Function for C$ button
Function Get-Cdrive {
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    # Clear out old PSDrive if it exists. This custom drive is created in this script for use with C$.
    Remove-PSDrive -Name "TempCommandCenter" -Force -ErrorAction SilentlyContinue

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try connecting to remote computer's C$ based on quser results
    If ($Proceed -eq $true) {
        
        # Create PS Drive mapped to the computer's C:\Windows\CCM\Logs folder
        Try{
            New-PSDrive -Name "TempCommandCenter" -PSProvider "FileSystem" -Root "\\$Computer\C$" -Credential $LAcreds -ErrorAction Stop
    
            # Open the PS Drive
            Set-Location TempCommandCenter:

            # Open File Explorer
            Invoke-Item .

            # Display info in the info box
            $Output1 = [pscustomobject]$Status
            Update-InfoBox -Status "$Output1"

        }Catch{ # Catch most likely error, bad password
            Write-Host "Unable to start C$, check LA account password"

            # Display info in the info box
           Update-InfoBox -Status "Unable to start C$, check LA account password"          
        }
    }ElseIf($Proceed -eq $False){

        # Display info in the info box
        Update-InfoBox -Status 'Unable to connect to computer. Possible reasons: Computer offline, DNS problems, misspelled name'
    }
}

# Function to launch Powershell with LA account
Function Open-PowershellLA{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear out the previous job (if applicable)
    Remove-Job -Name PowershellLA -ErrorAction SilentlyContinue
    
    # Clear info boxes
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue
    
    # Open Powershell using LA creds
    $PSLAJob = Start-Job -ScriptBlock {Start-Process C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Verb RunAs} -Credential $LAcreds -Name PowershellLA
        
    # Start the job
    $PSLAJob
    
    # Verify the job went through correctly. Display message if not
    $Verify = $PSLAJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    If ($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.') {
            
        Write-Host "Unable to start Powershell LA due to bad password"
            
        Update-InfoBox -Status "Unable to open Powershell with LA account. Bad password"
    }Else {
        Write-Host "Running Powershell using LA account"
        Update-InfoBox -Status "Launching Powershell with LA account..."
    }
}

# Function to RDP using mstsc
Function Start-RDP{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info boxes
    $WPFInfoBox1.Clear()
    
    # Get computer
    $Computer = $WPFComputerNameBox.Text
    
    # Verify computer is online, if so start RDP
    Try {
    
        $PingTest = Test-Connection -ComputerName $Computer -Count 1 
    
            If ($PingTest) {
    
                Write-Host "Ping successful, attempting to RDP"    
                Update-InfoBox -Status "Attempting to RDP to $Computer."
                    
                # Open RDP with connection to $Computer
                Start-Process Mstsc.exe /v:$Computer

            }Else{
    
                Write-Host "Ping failed, unable to start RDP"
    
                Update-InfoBox -Status "Unable to RDP to $Computer. Verify it is online."
            }
    }Catch{
    
            Write-Host "Unable to start RDP for other reasons"    
            Update-InfoBox -Status "Unable to RDP to $Computer"
    } 
}

# Function to Ping computer
Function Start-Ping{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear the info box
    $WPFInfoBox1.Clear()
    
    # Get computer
    $Computer = $WPFComputerNameBox.Text
    
        # Ping Computer
            Try { 
                $PingTest = Test-Connection -ComputerName $Computer -Count 1 
                    If ($PingTest) {
                            
                        Write-Host "Ping successful"
                        $Address = $PingTest.ProtocolAddress | Out-String    
                        Update-InfoBox -Status "Ping successful, response from $Address"

                    }Else{    
                        Write-Host "Ping failed"    
                        Update-InfoBox -Status "Ping failed"
                    } 
            }Catch{
                Write-Host "Ping failed"
                Update-InfoBox -Status "Ping failed"    
            } 
}

# Function to launch MECM Console with LA account (admin rights)
Function Start-MECMConsole{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Verify MECM console is installed
    $MECMConsoleEXE = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"
    $ConsoleInstalled = Test-Path -Path $MECMConsoleEXE

    # MECM console is installed. Open it.
    If ($ConsoleInstalled -eq $true) {

        # Launch MECM console with LA creds
        Try {
            Start-Process -FilePath $MECMConsoleEXE -Credential $LAcreds -ErrorAction Stop
            Update-InfoBox -Status "Launching MECM Console with LA account..."

        }Catch{ # Catch most likely error, bad password
            # Unable to open console due to bad password
            Write-Host "Unable to open MECM console due to bad password"
            Update-InfoBox -Status "Unable to open MECM console, check LA account password."
        }
    }ElseIf($ConsoleInstalled -eq $false) { # MECM console not installed. Display error message.
            
        Write-Host "MECM console doesnt appear to be installed"
        
        # Display message in infobox
        Update-InfoBox -Status "MECM console is not installed..."
    }
}

# Function to launch SSMS with LA account (admin rights)
Function Start-SSMS{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Verify SSMS is installed
    $SSMSEXE = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"
    $SSMSInstalled = Test-Path -Path $SSMSEXE

    # SSMS is installed. Open it
    If ($SSMSInstalled -eq $true) {
       
        Try {
            Start-Process -FilePath $SSMSEXE -Credential $LAcreds -ErrorAction Stop

        }Catch{ # Catch most likely error, bad password

            # Unable to open SSMS due to bad password
            Write-Host "Unable to open SSMS due to bad password"
            Update-InfoBox -Status "Unable to open SSMS due to bad password..."
        }
    }ElseIf($SSMSInstalled -eq $False) { # SSMS not installed, display error message
        
        Write-Host "SQL Server Mgmt Studio doesnt appear to be installed"

        # Display message in infobox
        Update-InfoBox -Status "SQL Server Mgmt Studio is not installed"
    }
}

# Function to launch the Hyper-V Console with LA account (admin rights)
Function Start-HyperVConsole{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue

    $HyperVJob = Start-Job -ScriptBlock {Start-Process C:\Windows\System32\mmc.exe -ArgumentList C:\Windows\System32\virtmgmt.msc} -Credential $LAcreds -ErrorAction Stop

    # Start the job
    Update-InfoBox -Status "Launching Hyper-V Console..."
    $HyperVJob
    
    # Verify the job went through correctly. Display message if not
    $Verify = $HyperVJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    If ($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.') {
            
        Write-Host "Unable to start Hyper-V Console due to bad password"            
        Update-InfoBox -Status "Unable to start Hyper-V Console. Bad password..."
    }Else{

        Write-Host "Running Hyper-V Console using LA account"
        Update-InfoBox -Status "Running Hyper-V Console using LA account..."
    }
}

# Function to launch the Visual Studio Code with LA account (admin rights)
Function Start-VSCode{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue

    $VSCodeEXE = "C:\Program Files\Microsoft VS Code\Code.exe"
    $VSCodeJob = Start-Job -ScriptBlock { Start-Process -Path $VSCodeEXE } -Credential $LAcreds -ErrorAction Stop

    # Start the job
    Update-InfoBox -Status "Launching Visual Studio Code with LA account..."
    $VSCodeJob
    
    # Verify the job went through correctly. Display message if not
    $Verify = $VSCodeJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    If ($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.') {
            
        Write-Host "Unable to start Visual Studio Code due to bad password"            
        Update-InfoBox -Status "Unable to start Visual Studio Code. Bad password..."
    }Else{

        Write-Host "Running Visual Studio Code using LA account"
        Update-InfoBox -Status "Running Visual Studio Code using LA account..."
    }
}

# Function to launch GPMC with LA account (admin rights)
Function Start-GroupPolicy{
    
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue

    $GroupPolicyJob = Start-Job -ScriptBlock {Start-Process "C:\Windows\System32\mmc.exe" -ArgumentList "C:\Windows\System32\gpmc.msc"} -Credential $LAcreds -ErrorAction Stop

    # Start the job
    $GroupPolicyJob
    
    # Verify the job went through correctly. Display message if not
    $Verify = $GroupPolicyJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    If ($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.') {
            
        Write-Host "Unable to start Group Policy Mgmt due to bad password"
            
        Update-InfoBox -Status "Unable to start Group Policy Mgmt. Bad password..."
    }Else{
        Write-Host "Running Group Policy Mgmt using LA account"
        Update-InfoBox -Status "Running Group Policy Mgmt using LA account..."
    }
}

################
# Function 8
# Create function for Query User button
################
Function Custom-QueryUser {
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear the info boxes
    $WPFInfoBox1.Clear()

   # Get computer
   $Computer = $WPFComputerNameBox.Text
    
   # Test if computer is online
   $PingTest = Test-Connection -ComputerName $Computer -Count 1
    
       If ($PingTest) {
            $Query = quser /server:$Computer 2>&1
            If ($LASTEXITCODE -ne 0) {
                $ErrorDetail = $Query.Exception.Message[1]
                Switch -Wildcard ($Query) {
                    '*[1722]*' { 
                        $Status = 'Remote RPC not enabled'
                    }
                    '*[5]*' {
                        $Status = 'Access denied'
                    }
                    'No User exists for*' {
                        $Status = 'No user logged on'
                    }
                    default {
                        $Status = 'Error'
                        $ErrorDetail = $Query.Exception.Message
                    }
                }
                $Output1 = [pscustomobject]$Status
                $Output2 = [pscustomobject]$ErrorDetail
                $WPFInfoBox1.AddText("$Output1")
                $WPFInfoBox2.AddText("$Output2")
            }Else{ 
                Update-InfoBox -Status "$Query"
            }
        }Else{
            Update-InfoBox -Status "Computer offline or other error..."
        }
}

# Function to launch Computer Management of remote computer
Function Start-CompMgmt{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Get computer
    $Computer = $WPFComputerNameBox.Text

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    If ($Proceed -eq $True) {
        
        Try { # Launch computer management and connect to remote computer
            $PowershellEXE = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
            Start-Process $PowershellEXE -Credential $LAcreds -ArgumentList "C:\Windows\System32\Compmgmt.msc /Computer:$Computer" -ErrorAction Stop

        }Catch{ # Most likely reason Comp Mgmt didnt open, bad password
            Write-Host "Unable to open Comp Mgmt, bad password"
            Update-InfoBox -Status "Unable to open Comp Mgmt. Check LA account password..."
        }

    }ElseIf($Proceed -eq $False) { # Verify-Computer failed, dont try to open Comp Mgmt
        
        Update-InfoBox -Status "Unable to connect to $Computer..."
    }
}

# Function to launch Print Management with LA account (admin rights)
Function Start-PrintMgmt{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue

    # Launch print management
    $PrintMgmtJob = Start-Job -ScriptBlock {Start-Process Printmanagement.msc} -Credential $LAcreds

    # Start the job
    $PrintMgmtJob

    # Verify the job went through correctly. Display message if not
    $Verify = $PrintMgmtJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    # Most likely reason Print Mgmt doesnt open, bad password
    If ($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.'){
                
        Write-Host "Unable to start Print Mgmt due to bad password"                
        Update-InfoBox -Status "Unable to start Print Mgmt. Bad password..."
    }Else{
            Write-Host "Running Print Mgmt using LA account"
            Update-InfoBox -Status "Running Print Mgmt using LA account..."
    }
}

# Function to launch Remote CMD (psexec)
Function Start-RemoteCMD{
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try Starting remote CMD using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If ($Proceed -eq $True) {
                
        Try { # Start remote CMD using PSExec
            # Start PSExec 
            $PSEXECEXE = "C:\ProgramData\ACG\CommandCenter\PsExec.exe"
            Start-Process -FilePath $PSEXECEXE -Credential $LAcreds -ArgumentList "\\$Computer cmd"

            Update-InfoBox -Status "Starting remote CMD on $Computer..."

        }Catch{ # Catch most likely error, bad password
            # Bad password 
            Write-Host "Unable to start remote CMD, check LA account password"

            # Display info in the info box
            Update-InfoBox -Status "Unable to start remote CMD, check LA account password..."
        }
    }ElseIf($Proceed -eq $False){ # $Proceed indicates problem continuing

        Update-InfoBox -Status 'Unable to start remote CMD. Possible reasons: Computer offline, DNS problems, misspelled name'
    }
}

# Create function for Info button
Function Get-ActiveDirectoryOU{
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:

    # Clear variables
    Clear-Variable -Name ADOU -ErrorAction SilentlyContinue
    
    # Clear info box
    $WPFInfoBox1.Clear()

    Try {
        $ADOU = Get-ADComputer -Identity $Computer -Properties * | Select-Object -ExpandProperty CanonicalName -ErrorAction Stop
        Update-InfoBox -Status "$ADOU"

    }Catch{
        Update-InfoBox -Status "Could not get AD info for $Computer"
    }

}

################
# Function 13
# Create function for Remote Control button
################
Function Start-RemoteControl {
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try Starting remote CMD using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If ($Proceed -eq $True) {
        
        # Verify MECM installed
        $MECMInstalled = Test-Path -Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"

        # Start remote control
        If ($MECMInstalled -eq $true) {
            Try {
                Start-Process -FilePath "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe" -ArgumentList "$Computer" -Credential $LAcreds -ErrorAction Stop
                $WPFInfoBox1.AddText("Starting remote control on $Computer")
            }
    
            Catch {
                $WPFInfoBox1.AddText("Could not start remote control on $Computer")
            }
        }
    
        ElseIf ($MECMInstalled -eq $false) {
            $WPFInfoBox1.AddText("Cannot find MECM remote control. Verify MECM console installed")
        }
    }

    # $Proceed indicates problem continuing
    ElseIf ($Proceed -eq $False) {

        $Status = 'Unable to start remote CMD. Possible reasons: Computer offline, DNS problems, misspelled name'
        
        # Display info in the info box
        $Output1 = [pscustomobject]$Status
        $WPFInfoBox1.AddText("$Output1")
    }
}

################
# Function 14
# Run GPUpdate script from MECM
################
function GP-Update {
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Clear info box
    $WPFInfoBox1.Clear()
    Try {
        Set-Location 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin'
        Import-Module .\ConfigurationManager.psd1 -ErrorAction Stop
        #!!!!!!!!# Update New-PSDrive command to contain 3 letter Site Code and FQDN of MP
        New-PSDrive -Name "XXX" -PSProvider "CMSite" -Root "FQDNofMP" -Description "Primary site" -ErrorAction Stop
        Set-Location ACM: -ErrorAction Stop
    }

    Catch {
        $WPFInfoBox1.AddText("Unable to load MECM ")
    }


    $Device = Get-CMDevice -Name $Computer

    If ($Device.CNIsOnline -eq $True) {
    
        Write-Host "Device is online"

        Try {
            #!!!!!!!!# This function runs a Script that is already created in MECM\Software Library\Scripts. Update ScriptGUID
            Invoke-CMScript -Device $Device -ScriptGuid 'SCRIPT-GUID-GOES-HERE' -ErrorAction Stop

            $WPFInfoBox1.AddText("Updating group policy on $Computer")
        }

        Catch {
            $WPFInfoBox1.AddText("Unable to update group policy on $Computer")
        }
    }

    ElseIf ($Device.CNIsOnline -eq $False) {
        Write-Host "Device is not online"
        $WPFInfoBox1.AddText("$Computer is not online")
    }

    Else {
        Write-Host "Cannot find device"
        $WPFInfoBox1.AddText("Cannot find $Computer")
    }


}

################
# Function 15
# Run machine policy update
################
function MECM-PolicyUpdate {
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Clear info box
    $WPFInfoBox1.Clear()

    Set-Location 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin'
    Import-Module .\ConfigurationManager.psd1
    #!!!!!!!!# Set-Location to 3 letter Site Code
    Set-Location XXX:

    $Device = Get-CMDevice -Name $Computer

    If ($Device.CNIsOnline -eq $True) {
    
        Write-Host "Device is online"

        Try {
            #!!!!!!!!# This forces a Machine Policy Retrieval on the client. There are other ways to do this, specifically the MECM-RunActions function (Found in EMTCustomModules.ps1) which will run ALL actions, not just this one.
            #!!!!!!!!# OR the Run-CMClientAction function found on line 1243 of this script
            Invoke-CMClientAction -DeviceName "$Computer" -ActionType ClientNotificationRequestMachinePolicyNow -ErrorAction Stop

            $WPFInfoBox1.AddText("Updating MECM policy on $Computer")
        }

        Catch {
            $WPFInfoBox1.AddText("Unable to update MECM policy on $Computer")
        }
    }

    ElseIf ($Device.CNIsOnline -eq $False) {
        Write-Host "Device is not online"
    }

}

################
# Function 16
# Create function for Remote Registry button
################
Function Start-RemoteRegistry{
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Clear variable
    Clear-Variable -Name CheckFile -ErrorAction SilentlyContinue

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try Starting remote CMD using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If ($Proceed -eq $True) {

        # Store computer name in a place accessible to the next script
        $CheckFile = Test-Path -Path C:\Users\Public\Documents\RemoteReg.txt

        If ($CheckFile -eq $False) {

            New-Item -Path C:\Users\Public\Documents\RemoteReg.txt
            Set-Content -Path C:\Users\Public\Documents\RemoteReg.txt -Value $Computer
        }

        ElseIf ($CheckFile -eq $True) {

            Set-Content -Path C:\Users\Public\Documents\RemoteReg.txt -Value $Computer
        }

        # Open Powershell using LA creds, then run keystrokes script
        $RemoteRegJob = Start-Job -ScriptBlock {Start-Process "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File C:\ProgramData\ACG\CommandCenter\RemoteRegistry.ps1"} -Credential $LAcreds -Name RemoteRegLA
        
        # Start the job
        $RemoteRegJob
        $WPFInfoBox1.AddText("Starting remote registry on $Computer, please wait")
    }

    # $Proceed indicates problem continuing
    ElseIf ($Proceed -eq $False) {

        $Status = 'Unable to start remote CMD. Possible reasons: Computer offline, DNS problems, misspelled name'
        
        # Display info in the info box
        $Output1 = [pscustomobject]$Status
        $WPFInfoBox1.AddText("$Output1")
    }
}

################
# Function 17
# Create function for Start Remote Registry button
################
Function Enable-RemoteRegistry {
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try starting service using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If ($Proceed -eq $True) {
        
        # Enable remote registry using PSExec
        Try {
            # Start PSExec 
            Start-Process -FilePath "C:\ProgramData\ACG\CommandCenter\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer Sc config remoteregistry start=demand"
            Start-Process -FilePath "C:\ProgramData\ACG\CommandCenter\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer net start remoteregistry"

            $WPFInfoBox1.AddText("Started remote registry service on $Computer")
            }

        # Catch most likely error, bad password
        Catch {
            # Bad password 
            Write-Host "Unable to start remote registry service, check LA account password"

            # Display info in the info box
            $WPFInfoBox1.AddText("Unable to start remote registry service, check LA account password")
            }
        }
    
    # $Proceed indicates problem continuing
    ElseIf ($Proceed -eq $False) {

        $Status = 'Unable to start remote registry service. Possible reasons: Computer offline, DNS problems, misspelled name'
        
        # Display info in the info box
        $Output1 = [pscustomobject]$Status
        $WPFInfoBox1.AddText("$Output1")
        }
}

################
# Function 18
# Create function for Stop Remote Registry button
################
Function Disable-RemoteRegistry {
    param (
        $Computer = $WPFComputerNameBox.Text
    )

    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try stopping service using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If ($Proceed -eq $True) {
        
        # Disable remote registry using PSExec
        Try {
            # Start PSExec 
            Start-Process -FilePath "C:\ProgramData\ACG\CommandCenter\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer net stop remoteregistry"
            Start-Process -FilePath "C:\ProgramData\ACG\CommandCenter\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer Sc config remoteregistry start=disabled"

            $WPFInfoBox1.AddText("Stopped remote registry service on $Computer")
            }

        # Catch most likely error, bad password
        Catch {
            # Bad password 
            Write-Host "Unable to stop remote registry service, check LA account password"

            # Display info in the info box
            $WPFInfoBox1.AddText("Unable to stop remote registry service, check LA account password")
            }
        }
    
    # $Proceed indicates problem continuing
    ElseIf ($Proceed -eq $False) {

        $Status = 'Unable to start remote registry service. Possible reasons: Computer offline, DNS problems, misspelled name'
        
        # Display info in the info box
        $Output1 = [pscustomobject]$Status
        $WPFInfoBox1.AddText("$Output1")
        }
}

# Function to copy the CMClient Files to the Remote computer
Function Copy-CMClientFiles{
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try copying the MECM Client files to the C$ of the client computer based on quser results
    # $ProceedLogs indicates to continue
    If($Proceed -eq $True){
        Try{
            # Copy CMClient files to the remote system
            #!!!!!!!!# This copies the CMClient Folder to C:\Temp\CMClient on the local host. If you want it in a different location, change here.
            Update-InfoBox -Status "Copying CMClient files..."
            $Temp = "\\$Computer\C$\Temp"
            $Temp_CMClient_Loc = "\\$Computer\C$\Temp\CMClient"

            if(!(Test-Path -Path "$Temp")){
                Write-Host "$Temp does not exist on $Computer. Creating..."
                Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s cmd mkdir c:\temp"            
            }
            
            If(Test-Path -Path "$Temp_CMClient_Loc\ccmsetup.exe"){
                Write-Host "C:\Temp\CMClient folder already exists on $Computer..."
                Update-InfoBox -Status "C:\Temp\CMClient folder already exists on $Computer..."
                #Remove-Item -Path "$Temp_CMClient_Loc" -Recurse -Force -Verbose
            }else{
                Write-Host 'Copying Latest and Greatest CMClient Files to C:\Temp\CMClient'
                Update-InfoBox -Status 'Copying Latest and Greatest CMClient Files to C:\Temp\CMClient'
                Copy-Item -Path "$PSScriptRoot\CMClient" -Destination $Temp_CMClient_Loc -Recurse -Force -Verbose
            }

            # Import EMT Custom Powershell Modules
            $PowershellModules = "\\$Computer\C$\Program Files\WindowsPowerShell\Modules"
            $PowershellModules_Local = "C:\Program Files\WindowsPowerShell\Modules"
            #!!!!!!!!# EMT Custom Modules are copied to the local machine and imported with the name via creating this folder. If you change the Name of the file, update here.
            if(!(Test-Path -Path $PowershellModules\EMTCustomModules)){
                New-Item -Path "$PowershellModules" -Name "EMTCustomModules" -ItemType Directory -Force
                Update-InfoBox -Status "Creating EMTCustomModules folder..."
            }

            Start-Sleep -seconds 5
            #!!!!!!!!# ALSO CHANGE HERE
            $CustomEMTModules = (Get-ChildItem -Path "$PSScriptRoot\CMClient" -filter "*.psm1").FullName
            Copy-Item -Path "$PSScriptRoot\CMClient\*.psm1" -Destination "$PowershellModules\EMTCustomModules" -Recurse -Force -Verbose
            Start-Sleep -Seconds 5
            #Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s powershell Copy-Item -Path $Temp_CMClient_Loc\EMTCustomModules.psm1 -Destination $PowershellModules_Local\EMTCustomModules -Recurse -force"


            

            
            Write-Host "Attempting to import custom powershell modules..."
            Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s powershell Set-Executionpolicy bypass -force"
            Start-Sleep -Seconds 5 
            #!!!!!!!!# ALSO CHANGE HERE
            Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s powershell Import-Module -Name EMTCustomModules -verbose -force"
 
            Update-InfoBox -Status "Custom EMT Powershell modules have been imported..."
            
        }Catch{
            Update-InfoBox -Status 'Unable to copy files to the client.'
        }
    }Else{
        Update-InfoBox -Status "Unable to copy the CMClient files to $Computer. Possible reasons: Computer offline, DNS problems, misspelled name."
    }
}

# Function to start a PSRemoting session with a Remote computer
Function Start-PSRemoting{
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Make sure powershell goes back to the C drive to avoid errors
    Set-Location C:
    
    # Clear info box
    $WPFInfoBox1.Clear()

    # Use Verify-Computer to test if computer is actually online. Helps verify DNS record is accurate as well.
    Write-Host "Running the Verify-Computer function"
    Verify-Computer
    Write-Host "Proceed is set to $Proceed"

    # Evaluate the $Proceed variable to see if we should try Starting PSRemoting using PSExec based on quser results
    # $ProceedLogs indicates to continue
    If($Proceed -eq $True){
        # Start PSRemoting using PSExec
        Try{
            # Start PSExec 
            Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s powershell Enable-PSRemoting -Force"

            Update-InfoBox -Status "Starting PSRemoting on $Computer"

            #Initiate the PSSession with the remote computer
            Enter-PSSession -ComputerName $Computer -Credential $LAcreds
        }Catch{ # Catch most likely error, bad password
            # Bad password 
            Write-Host "Unable to start PSRemoting on the client, check LA account password"

            # Display info in the info box
            Update-InfoBox -Status 'Unable to start PSRemoting on the Client, check LA account password'
        }
    }Else{

        Update-InfoBox -Status 'Unable to PSRemote to the client. Possible reasons: Computer offline, DNS problems, misspelled name'
    }
}

# Function to run MECM Client Actions
Function Run-CMClientAction{
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

# Function to uninstall the existing CMClient. If trying to run on a remote computer, you must use Copy-CMClientFiles first.
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
        #!!!!!!!!# Modify if you don't want all of these folders / files deleted.
        $FoldersArr = @("C:\Windows\ccm","C:\Windows\ccmsetup","C:\Windows\ccmcache","C:\Windows\ccmtemp","C:\SMS","C:\SMSTSLog",
                        "C:\_SMSTaskSequence","C:\Windows\Temp","C:\Windows\sms*.ini","C:\Windows\sms*.mif")

        Foreach($item in $FoldersArr){
            if(Test-Path $item){
                Remove-Item -Path $item -Recurse -Force -Verbose -ErrorAction SilentlyContinue 
                Write-Host "Removing folders..."                   
            }
        }

        #Deleting Reg Keys
        #!!!!!!!!# Modify if you don't want all of these Reg Keys deleted
        $RegKeysArr = @("HKLM:\Software\Microsoft\ccm","HKLM:\Software\Microsoft\CCMSETUP","HKLM:\Software\Microsoft\SMS",
                        "HKLM:\SYSTEM\CurrentControlSet\Services\CCMSetup","HKLM:\SYSTEM\CurrentControlSet\Services\CcmExec",
                        "HKLM:\SYSTEM\CurrentControlSet\Services\smstsmgr","HKLM:\SYSTEM\CurrentControlSet\Services\CmRcService")

        Foreach($regK in $RegKeysArr){
            if(Test-Path -Path $regK){
                Remove-Item -Path $regK -Recurse -Force -Verbose -ErrorAction SilentlyContinue
                Write-Host "Deleting CMClient Reg Keys..."                 
            }
        }

        #!!!!!!!!# This is the part where it does the same thing as WMIFix01.bat, but in powershell-y ways.
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

        #Done
        Write-Host "CMClient Uninstall Complete!"

        Update-InfoBox -Status "MECM Client has been uninstalled..."
        
    }else{ # if the CMClient is installed, nothing needs to happen.
        #Done
        Write-Host "CMClient is not installed"

        #Update-InfoBox -Status "MECM Client is not currently installed..."

    }


}

# Function to install the CMClient from C:\Temp\CMClient. If trying to run on a remote computer, you must use Copy-CMClientFiles first. 
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
        #!!!!!!!!# Update for your environment. If you don't have a CMG, # out lines 1454 and 1458
        #!!!!!!!!# If you do have a CMG and need help finding this info, try this website: https://eskonr.com/2020/05/how-to-prepare-sccm-cmg-client-installation-switches-for-internet-based-client/
        $CMGArguments = "/nocrlcheck /mp:HTTPS:// CCMHTTPSSTATE=31 CCMHOSTNAME= SMSSiteCode=XXX AADTENANTID= AADCLIENTAPPID= AADRESOURCEURI=api://"
        #!!!!!!!!# Modify to your install command and take out any switches that don't apply (PKI/CRL)
        $arguments = "/mp:FQDNofMP /logon SMSSITECODE=XXX SMSCACHESIZE=20000 /UsePKICert /NoCRLCheck /forceinstall"
        Start-Process -FilePath "C:\Temp\CMClient\ccmsetup.exe" -ArgumentList $arguments -Wait
        Start-Process -FilePath "C:\Temp\CMClient\ccmsetup.exe" -ArgumentList $CMGArguments -Wait
    
        #!!!!!!!!# This section will run ALL MECM actions on the client every 15 seconds until there are 2 items in c:\windows\ccmcache. Adjust to your liking.
        $Count = 1
        $ccmcache = 'C:\windows\ccmcache'
        Do{
            Run-CMClientAction -ClientAction AppDeployment
            Run-CMClientAction -ClientAction MachinePolicy
            Run-CMClientAction -ClientAction DiscoveryData
            Run-CMClientAction -ClientAction ComplianceEvaluation
            Run-CMClientAction -ClientAction HardwareInventory
            Run-CMClientAction -ClientAction UpdateDeployment
            Run-CMClientAction -ClientAction UpdateScan
            Run-CMClientAction -ClientAction SoftwareInventory
            Write-Output "MECM actions have run $Count times."
            Start-Sleep -Seconds 15
            $Count ++
    
        }While((Get-ChildItem -Path "$ccmcache").count -lt 2)
    
    }
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

# Function to open Windows Terminal with LA account (admin rights)
Function Open-TerminalLA{
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Clear out the previous job (if applicable)
    Remove-Job -Name TerminalLA -ErrorAction SilentlyContinue
    
    # Clear info boxes
    $WPFInfoBox1.Clear()

    # Clear common variable used in multiple functions
    Clear-Variable -Name Verify -ErrorAction SilentlyContinue
    
    # Open Powershell using LA creds
        # Gets the EXE regardless of version
    $TerminalEXE = (Get-ChildItem -Path "C:\Program Files\WindowsApps" -filter "WindowsTerminal.exe" -depth 2).Fullname
    $WTLAJob = Start-Job -ScriptBlock {Start-Process $TerminalEXE -Verb RunAs} -Credential $LAcreds -Name TerminalLA
        
    # Start the job
    $WTLAJob
    
    # Verify the job went through correctly. Display message if not
    $Verify = $WTLAJob.ChildJobs[0].JobStateInfo.Reason.Message
    
    If($Verify -eq 'An error occurred while starting the background process. Error reported: The user name or password is incorrect.'){
            
        Write-Host "Unable to start Terminal LA due to bad password"            
        Update-InfoBox -Status "Unable to open Terminal with LA account. Bad password"

    }Else{
        Write-Host "Running Terminal using LA account"
        Update-InfoBox -Status "Launching Windows Terminal with LA account..."
    }
}

# Function to Remove Computer from AD
function Remove-ComputerAD{
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Clear info box
    $WPFInfoBox1.Clear()

    Import-Module ActiveDirectory

    $ADDevice = Get-ADComputer -Identity $Computer

    if($null -ne $ADDevice){
        Write-Host "Found $Computer in Active Directory. Removing..."
        Update-InfoBox -Status "Found $Computer in Active Directory. Removing..."
        $ADDevice | Remove-ADObject -Recursive -Verbose -Confirm:$false

        Start-Sleep -Seconds 5
        $ADDevice2 = Get-ADComputer -Identity $Computer
        if($null -eq $ADDevice2){
            Write-Host "$Computer was successfully removed from Active Directory."
            Update-InfoBox -Status "$Computer was successfully removed from Active Directory."
        }else{
            Write-Host "$Computer still detected in Active Directory after attempted removal..."
            Update-InfoBox -Status "$Computer still detected in Active Directory after attempted removal..."
        }
    }else{
        Write-Host "Unable to find $Computer in Active Directory..."
        Update-InfoBox -Status "Unable to find $Computer in Active Directory..."
    }

}

# Function to Remove Computer from MECM
function Remove-ComputerMECM{
    param (
        $Computer = $WPFComputerNameBox.Text
    )
    
    # Clear info box
    $WPFInfoBox1.Clear()

   # Set-Location 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin'
   # Import-Module .\ConfigurationManager.psd1
   # Set-Location ACM:

    # Import the ConfigMgr Module
    Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1') | Out-Null
    $ReturnPath = (Get-Location).Path

    # Get the CMSite
    $Site = Get-PSDrive | Where-Object{$_.Provider -like '*CMSite'}

    # Mount the CMSite PS Drive
    Set-Location $Site':'

    $CMDevice = Get-CMDevice -Name "$Computer" | Select-Object -Property Name,ResourceID

    if($null -ne $CMDevice){
        Write-Host "Found $Computer in MECM. Removing..."
        Update-InfoBox -Status "Found $Computer in MECM. Removing..."
        Remove-CMResource -ResourceID $CMDevice.ResourceID -Force

    }else{
        Write-Host "Unable to find $Computer in MECM..."
        Update-InfoBox -Status "Unable to find $Computer in MECM..."

    }

    Set-Location $ReturnPath

}

#===========================================================================
# Buttons
#===========================================================================

################
# Program the C$ button
################
$WPFC_DriveButton.Add_Click({

    # Run function
    Get-CDrive
})

################
# Program the MECM Console button
################
$WPFMECM_Console_Button.Add_Click({
    
    # Run function
    Start-MECMConsole
})

################
# Program the Terminal button
################
$WPFTerminalButton.Add_Click({
    
    # Run function
    Open-TerminalLA
})

################
# Program the SQL Server Mgmt Studio button
################
$WPFSSMSButton.Add_Click({

    # Run function
    Start-SSMS
})

################
# Program the Group Policy button
################
$WPFGPO_Button.Add_Click({

    # Run function
    Start-GroupPolicy
})

################
# Program the Clear button
################
$WPFClearButton.Add_Click({
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    $WPFComputerNameBox.Clear()
    $WPFInfoBox1.Clear()
})

################
# Program the query user button
################
$WPFUserButton.Add_Click({  
      
    # Run function
    Custom-QueryUser
})

################
# Program the ping button
################
$WPFPingButton.Add_Click({

    # Run function
    Start-Ping
})

################
# Program the RDP button
################
$WPFRDPbutton.Add_Click({

    # Run function
    Start-RDP
})

################
# Program the Powershell LA button
################
$WPFPowershell_LA_Button.Add_Click({

    # Run function
    Open-PowershellLA
})

################
# Program the Powershell (normal) button
################
$WPFPowershellButton.Add_Click({
    # Make sure powershell goes back to the C drive to avoid errors caused by a missing TempCommandCenter PSDrive
    Set-Location C:

    # Open Powershell using normal creds
    Start-Job -ScriptBlock {Start-Process C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe}
})

################
# Program Computer Management button
################
#!!!!!!!!# Not included in this UI. UI can be modified.
$WPFCompMgmtButton.Add_Click({

    # Run function
    Start-CompMgmt
})

################
# Program Print Management button
################
$WPFPrintButton.Add_Click({
    
    # Run function
    Start-PrintMgmt
})

################
# Program Remote CMD button
################
#!!!!!!!!# Not included in this UI
$WPFRemoteCMDButton.Add_Click({
    
    # Run function
    Start-RemoteCMD
})

################
# Program Query AD OU button
################
$WPFQueryADOUButton.Add_Click({
    
    # Run function
    Get-ActiveDirectoryOU
})

################
# Program Remote Control button
################
#!!!!!!!!# Not included in this UI
$WPFRemoteControlButton.Add_Click({
    
    # Run function
    Start-RemoteControl
})

################
# Program GP Update button
################
#!!!!!!!!# Not included in this UI
$WPFGPUpdateButton.Add_Click({
    
    # Run function
    GP-Update
})

################
# Program Run All Actions button
################
$WPFMECMAllActionsButton.Add_Click({
    
    # Run functions
    $Computer = $WPFComputerNameBox.Text

    Copy-CMClientFiles
    Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer -s powershell MECM-RunActions" -Wait
    Update-InfoBox -Status "Running MECM Client Actions on $Computer"
})

################
# Program Remote Registry button
################
#!!!!!!!!# Not included in this UI
$WPFRegistryButton.Add_Click({
    
    # Run function
    Start-RemoteRegistry
    })

################
# Program Start Remote Registry button
################
#!!!!!!!!# Not included in this UI
$WPFRegStartButton.Add_Click({
    
    # Run function
    Enable-RemoteRegistry
    })

################
# Program Stop Remote Registry button
################
#!!!!!!!!# Not included in this UI
$WPFRegStopButton.Add_Click({
    
    # Run function
    Disable-RemoteRegistry
    })

################
# Program Uninstall MECM Client button
################
$WPFUninstallMECMClientButton.Add_Click({
    
    # Run functions
    $Computer = $WPFComputerNameBox.Text

    Copy-CMClientFiles
    Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer powershell Uninstall-CMClient" -Wait
})

################
# Program Install MECM Client button
################
$WPFInstallMECMClientButton.Add_Click({
    
    # Run functions
    $Computer = $WPFComputerNameBox.Text

    Copy-CMClientFiles
    Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer powershell Install-CMClient" -Wait
    })

################
# Program Hyper-V button
################
$WPFHyperVButton.Add_Click({
    
    # Run function
    Start-HyperVConsole
})


################
# Program VSCode button
################
$WPFVSCodeButton.Add_Click({
    
    # Run function
    Start-VSCode
})


################
# Program Clear CM Cache button
################
$WPFClearMECMCacheButton.Add_Click({
    
    # Run functions
    $Computer = $WPFComputerNameBox.Text

    Copy-CMClientFiles
    Start-Process -FilePath "$PSScriptRoot\PsExec.exe" -Credential $LAcreds -ArgumentList "\\$Computer powershell Clear-CMCache" -Wait
})


################
# Program Remove From MECM button
################
$WPFRemoveFromMECMButton.Add_Click({

    Remove-ComputerMECM

})


################
# Program Remove From AD button
################
$WPFRemoveFromADButton.Add_Click({

    Remove-ComputerAD

})
#===========================================================================
# Show the form
#===========================================================================

$Form.ShowDialog() | out-null