<# 
.SYNOPSIS 
   vDiagram Horizon Visio Drawing Tool

.DESCRIPTION
   vDiagram Horizon Visio Drawing Tool

.NOTES 
   File Name	: vDiagram_Horizon_1.0.1.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 1.0.1

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data
		MS Visio

.CHANGE LOG
	- 09/12/2021 - v1.0.1
		Initial release
#>

#region ~~< Parameters >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Param `
	( `
		[Switch] $debug,
		[Switch] $logcapture,
		[Switch] $logdraw
	)
#endregion ~~< Parameters >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Admin Check >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( -NOT ( [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole( [Security.Principal.WindowsBuiltInRole] "Administrator" ) ) `
	{ `
		Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
		Break
	}
#endregion ~~< Admin Check >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms" )
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Drawing" )
[void][System.Reflection.Assembly]::LoadWithPartialName( "PresentationFramework" )
#endregion ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pre-PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_PowerCliModule
{ `
	$PowerCliCheck =  [System.Windows.Forms.MessageBox]::Show( "PowerCLI Module was not found. Would you like to install? Click 'Yes' to install and 'No' cancel.","Warning! Powershell Module is missing.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning )
	switch  ( $PowerCliCheck ) `
	{ `
		'Yes' 
		{ `
			Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
			Install-Module -Name VMware.PowerCLI -Scope AllUsers -AllowClobber
			Write-Host "[$DateTime] Installing Module" $VMwareModule.Name -ForegroundColor Green
			$PowerCliUpdate = Get-Module -Name VMware.PowerCLI -ListAvailable
			Write-Host "[$DateTime] VMware PowerCLI Module" $PowerCliUpdate.Version "is installed." -ForegroundColor Green
		}
		'No'
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Unable to proceed without the PowerCLI Module installed. Please run script again and select to install module." -ForegroundColor Red
			exit
		}
	}
}
#endregion ~~< Find_PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Install PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCli = Get-Module -Name VMware.PowerCLI -ListAvailable
$PowerCliLatest = Find-Module -Name VMware.PowerCLI -Repository PSGallery -ErrorAction Stop
if ( ( $PowerCli.Name ) -match ( "VMware.PowerCLI" ) ) `
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] VMware PowerCLI Module(s)" $PowerCli.Version " found on this machine." -ForegroundColor Yellow
	if ( ( $PowerCliLatest.Version ) -gt ( $PowerCli.Version[0] ) ) `
	{ `
		$PowerCliUpgrade =  [System.Windows.Forms.MessageBox]::Show( "PowerCLI Module is not the latest. Would you like to upgrade? Click 'Yes' to upgrade and 'No' cancel.","Warning! PowerCLI Module is not the latest.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Information )
		switch  ( $PowerCliUpgrade ) `
		{ `
			'Yes' 
			{ `
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] You elected to upgrade VMware PowerCLI Module to " $PowerCliLatest.Version -ForegroundColor Yellow
				$Modules = Get-InstalledModule -Name VMware.*

				foreach ( $Module in $Modules ) `
				{ `
					$VMwareModules = Get-InstalledModule -Name $Module.Name -AllVersions
					foreach ( $VMwareModule in $VMwareModules )
					{ `
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Get-Module $VMwareModule.Name -ListAvailable | Uninstall-Module -Force
						Write-Host "[$DateTime] Uninstalling Module" $VMwareModule.Name $VMwareModule.Version -ForegroundColor Yellow
					}
				}

				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Installing latest VMware PowerCLI Module" -ForegroundColor Green
				Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
				Install-Module -Name VMware.PowerCLI -Scope AllUsers -AllowClobber
				$PowerCliUpdate = Get-Module -Name VMware.PowerCLI -ListAvailable
				Write-Host "[$DateTime] VMware PowerCLI Module" $PowerCliUpdate.Version "is installed." -ForegroundColor Green
			}
			'No'
			{ `
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] You elected not to upgrade VMware PowerCLI Module. Current version is" $PowerCli.Version[0] -ForegroundColor Yellow
			}
		}
	}
} `
else `
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] VMware PowerCLI Module is not installed." -ForegroundColor Red
	Find_PowerCliModule 
}
#endregion ~~< Install PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Pre-PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Post-Constructor Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< About >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
$MyVer = "1.0.1"
$LastUpdated = "September 12, 2021"
$About = 
@"

	vDiagram Horizon $MyVer
	
	Contributors:	Tony Gonzalez
			Jason Hopkins
	
	Description:	vDiagram Horizon $MyVer - Based off of Alan Renouf's vDiagram
	
	Created:		September 12, 2021
	
	Last Updated:	$LastUpdated

"@
#endregion ~~< About >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TestShapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( ( Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) -or $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) ) `
{ `
	$TestShapes = [System.Environment]::GetFolderPath( 'MyDocuments' ) + "\My Shapes\vDiagram_Horizon_" + $MyVer + ".vssx"
	if ( -not ( Test-Path $TestShapes ) )
	{ `
		$CurrentLocation = Get-Location
		$UpdatedShapes = "$CurrentLocation" + "\vDiagram_Horizon_" + "$MyVer" + ".vssx"
		Copy-Item $UpdatedShapes $TestShapes
		Write-Host "Copying Shapes File to My Shapes"
	}
	$shpFile = "\vDiagram_Horizon_" + $MyVer + ".vssx"
}
#endregion ~~< TestShapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Set_WindowStyle >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Set_WindowStyle {
Param `
( `
    [ Parameter() ]
    [ ValidateSet( 'FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL' ) ]
    $Style = 'SHOW',
    [ Parameter() ]
    $MainWindowHandle = ( Get-Process -Id $pid ).MainWindowHandle
)
    $WindowStates = @{ `
        FORCEMINIMIZE   = 11; HIDE            = 0
        MAXIMIZE        = 3;  MINIMIZE        = 6
        RESTORE         = 9;  SHOW            = 5
        SHOWDEFAULT     = 10; SHOWMAXIMIZED   = 3
        SHOWMINIMIZED   = 2;  SHOWMINNOACTIVE = 7
        SHOWNA          = 8;  SHOWNOACTIVATE  = 4
        SHOWNORMAL      = 1
    }
    Write-Verbose ( "Set Window Style {1} on handle {0}" -f $MainWindowHandle, $( $WindowStates[ $style ] ) )

    $Win32ShowWindowAsync = Add-Type -MemberDefinition @"
    [DllImport("user32.dll")] 
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -name "Win32ShowWindowAsync" -namespace Win32Functions -passThru

    $Win32ShowWindowAsync::ShowWindowAsync( $MainWindowHandle, $WindowStates[ $Style ] ) | Out-Null
}
#Set_WindowStyle MINIMIZE
if( $debug -eq $true)
{
	$ErrorActionPreference = "Continue"
	$WarningPreference = "Continue"
	Set_WindowStyle SHOWDEFAULT
}
if( $debug -eq $false)
{
	$ErrorActionPreference = "SilentlyContinue"
	$WarningPreference = "SilentlyContinue"
	Set_WindowStyle FORCEMINIMIZE
}
#endregion ~~< Set_WindowStyle >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< About_Config >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function About_Config 
{ `
	$About

    # Add objects for About
    $AboutForm = New-Object System.Windows.Forms.Form
    $AboutTextBox = New-Object System.Windows.Forms.RichTextBox
    
    # About Form
    $AboutForm.Icon = $Icon
    $AboutForm.AutoScroll = $True
    $AboutForm.ClientSize = New-Object System.Drawing.Size(464,500)
    $AboutForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $AboutForm.Name = "About"
    $AboutForm.StartPosition = 1
    $AboutForm.Text = "About vDiagram Horizon $MyVer"
    
    $AboutTextBox.Anchor = 15
    $AboutTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
    $AboutTextBox.BorderStyle = 0
    $AboutTextBox.Font = "Tahoma"
    $AboutTextBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $AboutTextBox.Location = New-Object System.Drawing.Point(13,13)
    $AboutTextBox.Name = "AboutTextBox"
    $AboutTextBox.ReadOnly = $True
    $AboutTextBox.Size = New-Object System.Drawing.Size(440,500)
    $AboutTextBox.Text = $About
        
    $AboutForm.Controls.Add( $AboutTextBox )

    $AboutForm.Show() | Out-Null
}
#endregion ~~< About_Config >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Post-Constructor Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Form Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vDiagram = New-Object System.Windows.Forms.Form
$vDiagram.ClientSize = New-Object System.Drawing.Size(1008, 661)
$CurrentLocation = Get-Location
$Icon = "$CurrentLocation" + "\vDiagram.ico"
$vDiagram.Icon = $Icon
$vDiagram.Text = "vDiagram Horizon " + $MyVer 
$vDiagram.BackColor = [System.Drawing.Color]::DarkCyan

#region ~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$components = New-Object System.ComponentModel.Container
#endregion ~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu = New-Object System.Windows.Forms.MenuStrip
$MainMenu.Location = New-Object System.Drawing.Point(0, 0)
$MainMenu.Size = New-Object System.Drawing.Size(1010, 24)
$MainMenu.TabIndex = 0
$MainMenu.Text = "MainMenu"

#region ~~< File >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$FileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
$FileToolStripMenuItem.Text = "File"
#endregion ~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ExitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExitToolStripMenuItem.Size = New-Object System.Drawing.Size(92, 22)
$ExitToolStripMenuItem.Text = "Exit"
$ExitToolStripMenuItem.Add_Click({$vDiagram.Close()})
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@( $ExitToolStripMenuItem)))
#endregion ~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< File >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Help >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HelpToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$HelpToolStripMenuItem.Size = New-Object System.Drawing.Size(44, 20)
$HelpToolStripMenuItem.Text = "Help"
#endregion ~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutToolStripMenuItem.Size = New-Object System.Drawing.Size(107, 22)
$AboutToolStripMenuItem.Text = "About"
$AboutToolStripMenuItem.Add_Click({About_Config})
$HelpToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@( $AboutToolStripMenuItem)))
#endregion ~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Help >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$MainMenu.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@( $FileToolStripMenuItem, $HelpToolStripMenuItem)))
$vDiagram.Controls.Add( $MainMenu )

#endregion ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UpperTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UpperTabs = New-Object System.Windows.Forms.TabControl
$UpperTabs.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$UpperTabs.ItemSize = New-Object System.Drawing.Size(85, 20)
$UpperTabs.Location = New-Object System.Drawing.Point(10, 30)
$UpperTabs.Size = New-Object System.Drawing.Size(990, 98)
$UpperTabs.TabIndex = 1
$UpperTabs.Text = "UpperTabs"

#region ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Prerequisites = New-Object System.Windows.Forms.TabPage
$Prerequisites.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Prerequisites.Location = New-Object System.Drawing.Point(4, 24)
$Prerequisites.Padding = New-Object System.Windows.Forms.Padding(3)
$Prerequisites.Size = New-Object System.Drawing.Size(982, 70)
$Prerequisites.TabIndex = 0
$Prerequisites.Text = "Prerequisites"
$Prerequisites.ToolTipText = "Prerequisites: These items are needed in order to run this script."
$Prerequisites.BackColor = [System.Drawing.Color]::LightGray

#region ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellLabel = New-Object System.Windows.Forms.Label
$PowershellLabel.Location = New-Object System.Drawing.Point(10, 15)
$PowershellLabel.Size = New-Object System.Drawing.Size(75, 20)
$PowershellLabel.TabIndex = 0
$PowershellLabel.Text = "Powershell:"
$Prerequisites.Controls.Add( $PowershellLabel )
#endregion ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellInstalled = New-Object System.Windows.Forms.Label
$PowershellInstalled.Location = New-Object System.Drawing.Point(96, 15)
$PowershellInstalled.Size = New-Object System.Drawing.Size(350, 20)
$PowershellInstalled.TabIndex = 1
$PowershellInstalled.Text = ""
$Prerequisites.Controls.Add( $PowershellInstalled )
#endregion ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliLabel = New-Object System.Windows.Forms.Label
$PowerCliLabel.Location = New-Object System.Drawing.Point(450, 15)
$PowerCliLabel.Size = New-Object System.Drawing.Size(64, 20)
$PowerCliLabel.TabIndex = 4
$PowerCliLabel.Text = "PowerCLI:"
$Prerequisites.Controls.Add( $PowerCliLabel )
#endregion ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliInstalled = New-Object System.Windows.Forms.Label
$PowerCliInstalled.Location = New-Object System.Drawing.Point(520, 15)
$PowerCliInstalled.Size = New-Object System.Drawing.Size(400, 20)
$PowerCliInstalled.TabIndex = 5
$PowerCliInstalled.Text = ""
$Prerequisites.Controls.Add( $PowerCliInstalled )
#endregion ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleLabel = New-Object System.Windows.Forms.Label
$PowerCliModuleLabel.Location = New-Object System.Drawing.Point(10, 40)
$PowerCliModuleLabel.Size = New-Object System.Drawing.Size(110, 20)
$PowerCliModuleLabel.TabIndex = 2
$PowerCliModuleLabel.Text = "PowerCLI Module:"
$Prerequisites.Controls.Add( $PowerCliModuleLabel )
#endregion ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleInstalled = New-Object System.Windows.Forms.Label
$PowerCliModuleInstalled.Location = New-Object System.Drawing.Point(128, 40)
$PowerCliModuleInstalled.Size = New-Object System.Drawing.Size(320, 20)
$PowerCliModuleInstalled.TabIndex = 3
$PowerCliModuleInstalled.Text = ""
$Prerequisites.Controls.Add( $PowerCliModuleInstalled )
#endregion ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioLabel = New-Object System.Windows.Forms.Label
$VisioLabel.Location = New-Object System.Drawing.Point(450, 40)
$VisioLabel.Size = New-Object System.Drawing.Size(40, 20)
$VisioLabel.TabIndex = 6
$VisioLabel.Text = "Visio:"
$Prerequisites.Controls.Add( $VisioLabel )
#endregion ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioInstalled = New-Object System.Windows.Forms.Label
$VisioInstalled.Location = New-Object System.Drawing.Point(490, 40)
$VisioInstalled.Size = New-Object System.Drawing.Size(320, 20)
$VisioInstalled.TabIndex = 7
$VisioInstalled.Text = ""
$Prerequisites.Controls.Add( $VisioInstalled )
#endregion ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.Controls.Add( $Prerequisites )
#endregion ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServerInfo = New-Object System.Windows.Forms.TabPage
$ConnectionServerInfo.Location = New-Object System.Drawing.Point(4, 24)
$ConnectionServerInfo.Padding = New-Object System.Windows.Forms.Padding(3)
$ConnectionServerInfo.Size = New-Object System.Drawing.Size(982, 70)
$ConnectionServerInfo.TabIndex = 1
$ConnectionServerInfo.Text = "Connection Server Info"
$ConnectionServerInfo.BackColor = [System.Drawing.Color]::LightGray

#region ~~< ConnectionServerLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServerLabel = New-Object System.Windows.Forms.Label
$ConnectionServerLabel.Location = New-Object System.Drawing.Point(8, 11)
$ConnectionServerLabel.Size = New-Object System.Drawing.Size(120, 20)
$ConnectionServerLabel.TabIndex = 0
$ConnectionServerLabel.Text = "Connection Server:"
$ConnectionServerInfo.Controls.Add( $ConnectionServerLabel )
#endregion ~~< ConnectionServerLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnServTextBox = New-Object System.Windows.Forms.TextBox
$ConnServTextBox.Location = New-Object System.Drawing.Point(128, 8)
$ConnServTextBox.Size = New-Object System.Drawing.Size(188, 21)
$ConnServTextBox.TabIndex = 1
$ConnServTextBox.Text = ""
$ConnectionServerInfo.Controls.Add( $ConnServTextBox )
#endregion ~~< ConnectionServerTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServerToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$ConnectionServerToolTip.AutoPopDelay = 5000
$ConnectionServerToolTip.InitialDelay = 50
$ConnectionServerToolTip.IsBalloon = $true
$ConnectionServerToolTip.ReshowDelay = 100
$ConnectionServerToolTip.SetToolTip( $ConnServTextBox, "Enter Connection Server name" )
#endregion ~~< ConnectionServerToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameLabel = New-Object System.Windows.Forms.Label
$UserNameLabel.Location = New-Object System.Drawing.Point(324, 11)
$UserNameLabel.Size = New-Object System.Drawing.Size(70, 20)
$UserNameLabel.TabIndex = 2
$UserNameLabel.Text = "User Name:"
$ConnectionServerInfo.Controls.Add( $UserNameLabel )
#endregion ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameTextBox = New-Object System.Windows.Forms.TextBox
$UserNameTextBox.Location = New-Object System.Drawing.Point(402, 8)
$UserNameTextBox.Size = New-Object System.Drawing.Size(238, 21)
$UserNameTextBox.TabIndex = 3
$UserNameTextBox.Text = ""
$ConnectionServerInfo.Controls.Add( $UserNameTextBox )
#endregion ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$UserNameToolTip.AutoPopDelay = 5000
$UserNameToolTip.InitialDelay = 50
$UserNameToolTip.IsBalloon = $true
$UserNameToolTip.ReshowDelay = 100
$UserNameToolTip.SetToolTip( $UserNameTextBox, "Enter User Name."+[char]13+[char]10+[char]13+[char]10+"Example:"+[char]13+[char]10+"administrator@vsphere.local"+[char]13+[char]10+"Domain\User" )
#endregion ~~< UserNameToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Location = New-Object System.Drawing.Point(656, 11)
$PasswordLabel.Size = New-Object System.Drawing.Size(70, 20)
$PasswordLabel.TabIndex = 4
$PasswordLabel.Text = "Password:"
$ConnectionServerInfo.Controls.Add( $PasswordLabel )
#endregion ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Location = New-Object System.Drawing.Point(734, 8)
$PasswordTextBox.Size = New-Object System.Drawing.Size(238, 21)
$PasswordTextBox.TabIndex = 5
$PasswordTextBox.Text = ""
$PasswordTextBox.UseSystemPasswordChar = $true
$ConnectionServerInfo.Controls.Add( $PasswordTextBox )
#endregion ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$PasswordToolTip.AutoPopDelay = 5000
$PasswordToolTip.InitialDelay = 50
$PasswordToolTip.IsBalloon = $true
$PasswordToolTip.ReshowDelay = 100
$PasswordToolTip.SetToolTip( $PasswordTextBox, "Enter Passwrd."+[char]13+[char]10+[char]13+[char]10+"Characters will not be seen." )
#endregion ~~< PasswordToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$ConnectButton.Location = New-Object System.Drawing.Point(8, 37)
$ConnectButton.Size = New-Object System.Drawing.Size(345, 25)
$ConnectButton.TabIndex = 6
$ConnectButton.Text = "Connect to Horizon"
$ConnectButton.UseVisualStyleBackColor = $true
$ConnectionServerInfo.Controls.Add( $ConnectButton )
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$ConnectButtonToolTip.AutoPopDelay = 5000
$ConnectButtonToolTip.InitialDelay = 50
$ConnectButtonToolTip.IsBalloon = $true
$ConnectButtonToolTip.ReshowDelay = 100
$ConnectButtonToolTip.SetToolTip( $ConnectButton, "Click to connect to Horizon."+[char]13+[char]10+[char]13+[char]10+"If connected this button will turn green and show connected to the name entered in the Connection Server box."+[char]13+[char]10+[char]13+[char]10+"If disconnected or unable to connect this button will display red text, indicating that you were unable to"+[char]13+[char]10+"connect to Horizon either due to bad creditials, not being on the same network or insufficient access to Horizon." )
#endregion ~~< ConnectButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.Controls.Add( $ConnectionServerInfo )

#endregion ~~< ConnectionServerInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.SelectedIndex = 0
$vDiagram.Controls.Add( $UpperTabs )
#endregion ~~< UpperTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LowerTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LowerTabs = New-Object System.Windows.Forms.TabControl
$LowerTabs.Font = New-Object System.Drawing.Font( "Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$LowerTabs.Location = New-Object System.Drawing.Point(10, 136)
$LowerTabs.Size = New-Object System.Drawing.Size(990, 512)
$LowerTabs.TabIndex = 2
$LowerTabs.Text = "LowerTabs"

#region ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDirections = New-Object System.Windows.Forms.TabPage
$TabDirections.Location = New-Object System.Drawing.Point(4, 22)
$TabDirections.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDirections.Size = New-Object System.Drawing.Size(982, 486)
$TabDirections.TabIndex = 0
$TabDirections.Text = "Directions"
$TabDirections.UseVisualStyleBackColor = $true

#region ~~< Prerequisites Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesHeading = New-Object System.Windows.Forms.Label
$PrerequisitesHeading.Font = New-Object System.Drawing.Font( "Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$PrerequisitesHeading.Location = New-Object System.Drawing.Point(8, 8)
$PrerequisitesHeading.Size = New-Object System.Drawing.Size(149, 23)
$PrerequisitesHeading.TabIndex = 0
$PrerequisitesHeading.Text = "Prerequisites Tab"
$TabDirections.Controls.Add( $PrerequisitesHeading )
#endregion ~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesDirections = New-Object System.Windows.Forms.Label
$PrerequisitesDirections.Location = New-Object System.Drawing.Point(8, 32)
$PrerequisitesDirections.Size = New-Object System.Drawing.Size(900, 30)
$PrerequisitesDirections.TabIndex = 1
$PrerequisitesDirections.Text = "1. Verify that prerequisites are met on the "+[char]34+"Prerequisites"+[char]34+" tab."+[char]34+[char]13+[char]10+"2. If not please install needed requirements."+[char]13+[char]10
$TabDirections.Controls.Add( $PrerequisitesDirections )
#endregion ~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Prerequisites Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerInfo Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServerInfoHeading = New-Object System.Windows.Forms.Label
$ConnectionServerInfoHeading.Font = New-Object System.Drawing.Font( "Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$ConnectionServerInfoHeading.Location = New-Object System.Drawing.Point(8, 72)
$ConnectionServerInfoHeading.Size = New-Object System.Drawing.Size(250, 23)
$ConnectionServerInfoHeading.TabIndex = 2
$ConnectionServerInfoHeading.Text = "Connection Server Info Tab"
$TabDirections.Controls.Add( $ConnectionServerInfoHeading )
#endregion ~~< ConnectionServerInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServerInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServerInfoDirections = New-Object System.Windows.Forms.Label
$ConnectionServerInfoDirections.Location = New-Object System.Drawing.Point(8, 96)
$ConnectionServerInfoDirections.Size = New-Object System.Drawing.Size(900, 70)
$ConnectionServerInfoDirections.TabIndex = 3
$ConnectionServerInfoDirections.Text = "1. Click on"+[char]34+"ConnectionServer Info"+[char]34+" tab."+[char]13+[char]10+"2. Enter name of Connection Server"+[char]13+[char]10+"3. Enter Domain\User Name and Password (password will be hashed and not plain text)."+[char]13+[char]10+"4. Click on "+[char]34+"Connect to Connection Server"+[char]34+" button."
$TabDirections.Controls.Add( $ConnectionServerInfoDirections )
#endregion ~~< ConnectionServerInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ConnectionServerInfo Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvHeading = New-Object System.Windows.Forms.Label
$CaptureCsvHeading.Font = New-Object System.Drawing.Font( "Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$CaptureCsvHeading.Location = New-Object System.Drawing.Point(8, 176)
$CaptureCsvHeading.Size = New-Object System.Drawing.Size(216, 23)
$CaptureCsvHeading.TabIndex = 4
$CaptureCsvHeading.Text = "Capture CSVs for Visio Tab"
$TabDirections.Controls.Add( $CaptureCsvHeading )
#endregion ~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureDirections = New-Object System.Windows.Forms.Label
$CaptureDirections.Location = New-Object System.Drawing.Point(8, 200)
$CaptureDirections.Size = New-Object System.Drawing.Size(900, 65)
$CaptureDirections.TabIndex = 5
$CaptureDirections.Text = "1. Click on "+[char]34+"Capture CSVs for Visio"+[char]34+" tab."+[char]13+[char]10+"2. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select folder where you would like to output the CSVs to."+[char]13+[char]10+"3. Select items you wish to grab data on."+[char]13+[char]10+"4. Click on "+[char]34+"Collect CSV Data"+[char]34+" button."
$TabDirections.Controls.Add( $CaptureDirections )
#endregion ~~< CaptureDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Capture Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawHeading = New-Object System.Windows.Forms.Label
$DrawHeading.Font = New-Object System.Drawing.Font( "Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$DrawHeading.Location = New-Object System.Drawing.Point(8, 264)
$DrawHeading.Size = New-Object System.Drawing.Size(149, 23)
$DrawHeading.TabIndex = 6
$DrawHeading.Text = "Draw Visio Tab"
$TabDirections.Controls.Add( $DrawHeading )
#endregion ~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawDirections = New-Object System.Windows.Forms.Label
$DrawDirections.Location = New-Object System.Drawing.Point(8, 288)
$DrawDirections.Size = New-Object System.Drawing.Size(900, 130)
$DrawDirections.TabIndex = 7
$DrawDirections.Text = "1. Click on "+[char]34+"Select Input Folder"+[char]34+" button and select location where CSVs can be found."+[char]13+[char]10+"2. Click on "+[char]34+"Check for CSVs"+[char]34+" button to validate presence of required files."+[char]13+[char]10+"3. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select where location where you would like to save the Visio drawing."+[char]13+[char]10+"4. Select drawing that you would like to produce."+[char]13+[char]10+"5. Click on "+[char]34+"Draw Visio"+[char]34+" button."+[char]13+[char]10+"6. Click on "+[char]34+"Open Visio Drawing"+[char]34+" button once "+[char]34+"Draw Visio"+[char]34+" button says it has completed."
$TabDirections.Controls.Add( $DrawDirections )
#endregion ~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Draw Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add( $TabDirections )
#endregion ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCapture = New-Object System.Windows.Forms.TabPage
$TabCapture.Location = New-Object System.Drawing.Point(4, 22)
$TabCapture.Padding = New-Object System.Windows.Forms.Padding(3)
$TabCapture.Size = New-Object System.Drawing.Size(982, 486)
$TabCapture.TabIndex = 1
$TabCapture.Text = "Capture CSVs for Visio"
$TabCapture.UseVisualStyleBackColor = $true

#region ~~< TabCaptureToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCaptureToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$TabCaptureToolTip.AutoPopDelay = 5000
$TabCaptureToolTip.InitialDelay = 50
$TabCaptureToolTip.IsBalloon = $true
$TabCaptureToolTip.ReshowDelay = 100
$TabCaptureToolTip.SetToolTip( $TabCapture, "This must be ran first in order to collect the information"+[char]13+[char]10+"about your environment." )
#endregion ~~< TabCaptureToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Folder Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton = New-Object System.Windows.Forms.Button
$CaptureCsvOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCsvOutputButton.Location = New-Object System.Drawing.Point(220, 10)
$CaptureCsvOutputButton.Size = New-Object System.Drawing.Size(750, 25)
$CaptureCsvOutputButton.TabIndex = 1
$CaptureCsvOutputButton.Text = "Select Output Folder"
$CaptureCsvOutputButton.UseVisualStyleBackColor = $false
$CaptureCsvOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add( $CaptureCsvOutputButton )
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$CaptureCsvOutputButtonToolTip.AutoPopDelay = 5000
$CaptureCsvOutputButtonToolTip.InitialDelay = 50
$CaptureCsvOutputButtonToolTip.IsBalloon = $true
$CaptureCsvOutputButtonToolTip.ReshowDelay = 100
$CaptureCsvOutputButtonToolTip.SetToolTip( $CaptureCsvOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"CSV"+[char]39+"s."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green."+[char]13+[char]10+[char]13+[char]10+"If the folder has files in it you will be presented with an "+[char]13+[char]10+"option to move or delete the files that are currently there." )
#endregion ~~< CaptureCsvOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputLabel = New-Object System.Windows.Forms.Label
$CaptureCsvOutputLabel.Font = New-Object System.Drawing.Font( "Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$CaptureCsvOutputLabel.Location = New-Object System.Drawing.Point(10, 10)
$CaptureCsvOutputLabel.Size = New-Object System.Drawing.Size(210, 25)
$CaptureCsvOutputLabel.TabIndex = 0
$CaptureCsvOutputLabel.Text = "CSV Output Folder:"
$TabCapture.Controls.Add( $CaptureCsvOutputLabel )
#endregion ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$CaptureCsvBrowse.Description = "Select a directory"
$CaptureCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Capture Folder Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VirtualCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VirtualCenterCsvCheckBox.Checked = $true
$VirtualCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VirtualCenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$VirtualCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VirtualCenterCsvCheckBox.TabIndex = 2
$VirtualCenterCsvCheckBox.Text = "Export VirtualCenter Info"
$VirtualCenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $VirtualCenterCsvCheckBox )
#endregion ~~< VirtualCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VirtualCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$VirtualCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 40)
$VirtualCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VirtualCenterCsvValidationComplete.TabIndex = 3
$VirtualCenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $VirtualCenterCsvValidationComplete )
#endregion ~~< VirtualCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VirtualCenterCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$VirtualCenterCsvToolTip.AutoPopDelay = 5000
$VirtualCenterCsvToolTip.InitialDelay = 50
$VirtualCenterCsvToolTip.IsBalloon = $true
$VirtualCenterCsvToolTip.ReshowDelay = 100
$VirtualCenterCsvToolTip.SetToolTip( $VirtualCenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"VirtualCenters in Horizon." )
#endregion ~~< VirtualCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComposerServersCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ComposerServersCsvCheckBox.Checked = $true
$ComposerServersCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ComposerServersCsvCheckBox.Location = New-Object System.Drawing.Point(10, 60)
$ComposerServersCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ComposerServersCsvCheckBox.TabIndex = 4
$ComposerServersCsvCheckBox.Text = "Export Composer Server Info"
$ComposerServersCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $ComposerServersCsvCheckBox )
#endregion ~~< ComposerServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComposerServersCsvValidationComplete = New-Object System.Windows.Forms.Label
$ComposerServersCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 60)
$ComposerServersCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ComposerServersCsvValidationComplete.TabIndex = 5
$ComposerServersCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $ComposerServersCsvValidationComplete )
#endregion ~~< ComposerServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComposerServersCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$ComposerServersCsvToolTip.AutoPopDelay = 5000
$ComposerServersCsvToolTip.InitialDelay = 50
$ComposerServersCsvToolTip.IsBalloon = $true
$ComposerServersCsvToolTip.ReshowDelay = 100
$ComposerServersCsvToolTip.SetToolTip( $ComposerServersCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Composer Servers in Horizon." )
#endregion ~~< ComposerServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ComposerServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServersCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ConnectionServersCsvCheckBox.Checked = $true
$ConnectionServersCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ConnectionServersCsvCheckBox.Location = New-Object System.Drawing.Point(10, 80)
$ConnectionServersCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ConnectionServersCsvCheckBox.TabIndex = 6
$ConnectionServersCsvCheckBox.Text = "Export Connection Server Info"
$ConnectionServersCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $ConnectionServersCsvCheckBox )
#endregion ~~< ConnectionServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServersCsvValidationComplete = New-Object System.Windows.Forms.Label
$ConnectionServersCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 80)
$ConnectionServersCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ConnectionServersCsvValidationComplete.TabIndex = 7
$ConnectionServersCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $ConnectionServersCsvValidationComplete )
#endregion ~~< ConnectionServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServersCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$ConnectionServersCsvToolTip.AutoPopDelay = 5000
$ConnectionServersCsvToolTip.InitialDelay = 50
$ConnectionServersCsvToolTip.IsBalloon = $true
$ConnectionServersCsvToolTip.ReshowDelay = 100
$ConnectionServersCsvToolTip.SetToolTip( $ConnectionServersCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Connection Servers in Horizon." )
#endregion ~~< ConnectionServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ConnectionServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pools >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PoolsCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$PoolsCsvCheckBox.Checked = $true
$PoolsCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$PoolsCsvCheckBox.Location = New-Object System.Drawing.Point(10, 100)
$PoolsCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$PoolsCsvCheckBox.TabIndex = 8
$PoolsCsvCheckBox.Text = "Export Pool Info"
$PoolsCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $PoolsCsvCheckBox )
#endregion ~~< PoolsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PoolsCsvValidationComplete = New-Object System.Windows.Forms.Label
$PoolsCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 100)
$PoolsCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$PoolsCsvValidationComplete.TabIndex = 9
$PoolsCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $PoolsCsvValidationComplete )
#endregion ~~< PoolsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$VmHostCsvToolTip.AutoPopDelay = 5000
$VmHostCsvToolTip.InitialDelay = 50
$VmHostCsvToolTip.IsBalloon = $true
$VmHostCsvToolTip.ReshowDelay = 100
$VmHostCsvToolTip.SetToolTip( $PoolsCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Pools in Horizon." )
#endregion ~~< VmHostCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Pools >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Desktops >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DesktopsCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DesktopsCsvCheckBox.Checked = $true
$DesktopsCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DesktopsCsvCheckBox.Location = New-Object System.Drawing.Point(10, 120)
$DesktopsCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DesktopsCsvCheckBox.TabIndex = 10
$DesktopsCsvCheckBox.Text = "Export Desktop Info"
$DesktopsCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $DesktopsCsvCheckBox )
#endregion ~~< DesktopsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DesktopsCsvValidationComplete = New-Object System.Windows.Forms.Label
$DesktopsCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 120)
$DesktopsCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DesktopsCsvValidationComplete.TabIndex = 11
$DesktopsCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $DesktopsCsvValidationComplete )
#endregion ~~< DesktopsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DesktopsCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$DesktopsCsvToolTip.AutoPopDelay = 5000
$DesktopsCsvToolTip.InitialDelay = 50
$DesktopsCsvToolTip.IsBalloon = $true
$DesktopsCsvToolTip.ReshowDelay = 100
$DesktopsCsvToolTip.SetToolTip( $DesktopsCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Desktops in Horizon." )
#endregion ~~< DesktopsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Desktops >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RDSServersCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$RDSServersCsvCheckBox.Checked = $true
$RDSServersCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$RDSServersCsvCheckBox.Location = New-Object System.Drawing.Point(10, 140)
$RDSServersCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$RDSServersCsvCheckBox.TabIndex = 12
$RDSServersCsvCheckBox.Text = "Export RDS Server Info"
$RDSServersCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $RDSServersCsvCheckBox )
#endregion ~~< RDSServersCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RDSServersCsvValidationComplete = New-Object System.Windows.Forms.Label
$RDSServersCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 140)
$RDSServersCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$RDSServersCsvValidationComplete.TabIndex = 13
$RDSServersCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $RDSServersCsvValidationComplete )
#endregion ~~< RDSServersCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RDSServersCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$RDSServersCsvToolTip.AutoPopDelay = 5000
$RDSServersCsvToolTip.InitialDelay = 50
$RDSServersCsvToolTip.IsBalloon = $true
$RDSServersCsvToolTip.ReshowDelay = 100
$RDSServersCsvToolTip.SetToolTip( $RDSServersCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all RDS Servers in Horizon." )
#endregion ~~< RDSServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< RDSServers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farms >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FarmsCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$FarmsCsvCheckBox.Checked = $true
$FarmsCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FarmsCsvCheckBox.Location = New-Object System.Drawing.Point(10, 160)
$FarmsCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$FarmsCsvCheckBox.TabIndex = 14
$FarmsCsvCheckBox.Text = "Export Farms Info"
$FarmsCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $FarmsCsvCheckBox )
#endregion ~~< FarmsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FarmsCsvValidationComplete = New-Object System.Windows.Forms.Label
$FarmsCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 160)
$FarmsCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$FarmsCsvValidationComplete.TabIndex = 15
$FarmsCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $FarmsCsvValidationComplete )
#endregion ~~< FarmsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FarmsCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$FarmsCsvToolTip.AutoPopDelay = 5000
$FarmsCsvToolTip.InitialDelay = 50
$FarmsCsvToolTip.IsBalloon = $true
$FarmsCsvToolTip.ReshowDelay = 100
$FarmsCsvToolTip.SetToolTip( $FarmsCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Farms in Horizon." )
#endregion ~~< FarmsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Farms >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Applications >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ApplicationsCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ApplicationsCsvCheckBox.Checked = $true
$ApplicationsCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ApplicationsCsvCheckBox.Location = New-Object System.Drawing.Point(10, 180)
$ApplicationsCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ApplicationsCsvCheckBox.TabIndex = 16
$ApplicationsCsvCheckBox.Text = "Export Applications Info"
$ApplicationsCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $ApplicationsCsvCheckBox )
#endregion ~~< ApplicationsCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ApplicationsCsvValidationComplete = New-Object System.Windows.Forms.Label
$ApplicationsCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 180)
$ApplicationsCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ApplicationsCsvValidationComplete.TabIndex = 17
$ApplicationsCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $ApplicationsCsvValidationComplete )
#endregion ~~< ApplicationsCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ApplicationsCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$ApplicationsCsvToolTip.AutoPopDelay = 5000
$ApplicationsCsvToolTip.InitialDelay = 50
$ApplicationsCsvToolTip.IsBalloon = $true
$ApplicationsCsvToolTip.ReshowDelay = 100
$ApplicationsCsvToolTip.SetToolTip( $ApplicationsCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Applications in Horizon." )
#endregion ~~< ComposerServersCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Applications >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Gateways >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GatewaysCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$GatewaysCsvCheckBox.Checked = $true
$GatewaysCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$GatewaysCsvCheckBox.Location = New-Object System.Drawing.Point(310, 40)
$GatewaysCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$GatewaysCsvCheckBox.TabIndex = 18
$GatewaysCsvCheckBox.Text = "Export Gateways Info"
$GatewaysCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add( $GatewaysCsvCheckBox )
#endregion ~~< GatewaysCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GatewaysCsvValidationComplete = New-Object System.Windows.Forms.Label
$GatewaysCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 40)
$GatewaysCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$GatewaysCsvValidationComplete.TabIndex = 19
$GatewaysCsvValidationComplete.Text = ""
$TabCapture.Controls.Add( $GatewaysCsvValidationComplete )
#endregion ~~< GatewaysCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GatewaysCsvToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$GatewaysCsvToolTip.AutoPopDelay = 5000
$GatewaysCsvToolTip.InitialDelay = 50
$GatewaysCsvToolTip.IsBalloon = $true
$GatewaysCsvToolTip.ReshowDelay = 100
$GatewaysCsvToolTip.SetToolTip( $GatewaysCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Gateways in Horizon." )
#endregion ~~< GatewaysCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Gateways >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton = New-Object System.Windows.Forms.Button
$CaptureUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureUncheckButton.Location = New-Object System.Drawing.Point(8, 215)
$CaptureUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureUncheckButton.TabIndex = 50
$CaptureUncheckButton.Text = "Uncheck All"
$CaptureUncheckButton.UseVisualStyleBackColor = $false
$CaptureUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add( $CaptureUncheckButton )
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$CaptureUncheckButtonToolTip.AutoPopDelay = 5000
$CaptureUncheckButtonToolTip.InitialDelay = 50
$CaptureUncheckButtonToolTip.IsBalloon = $true
$CaptureUncheckButtonToolTip.ReshowDelay = 100
$CaptureUncheckButtonToolTip.SetToolTip( $CaptureUncheckButton, "Click to clear all check boxes above." )
#endregion ~~< CaptureUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton = New-Object System.Windows.Forms.Button
$CaptureCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCheckButton.Location = New-Object System.Drawing.Point(228, 215)
$CaptureCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureCheckButton.TabIndex = 51
$CaptureCheckButton.Text = "Check All"
$CaptureCheckButton.UseVisualStyleBackColor = $false
$CaptureCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add( $CaptureCheckButton )
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$CaptureCheckButtonToolTip.AutoPopDelay = 5000
$CaptureCheckButtonToolTip.InitialDelay = 50
$CaptureCheckButtonToolTip.IsBalloon = $true
$CaptureCheckButtonToolTip.ReshowDelay = 100
$CaptureCheckButtonToolTip.SetToolTip( $CaptureCheckButton, "Click to check all check boxes above." )
#endregion ~~< CaptureCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton = New-Object System.Windows.Forms.Button
$CaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureButton.Location = New-Object System.Drawing.Point(448, 215)
$CaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureButton.TabIndex = 52
$CaptureButton.Text = "Collect CSV Data"
$CaptureButton.UseVisualStyleBackColor = $false
$CaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add( $CaptureButton )
#endregion ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$CaptureButtonToolTip.AutoPopDelay = 5000
$CaptureButtonToolTip.InitialDelay = 50
$CaptureButtonToolTip.IsBalloon = $true
$CaptureButtonToolTip.ReshowDelay = 100
$CaptureButtonToolTip.SetToolTip( $CaptureButton, "Click to begin collecting environment information"+[char]13+[char]10+"on options selected above." )
#endregion ~~< CaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Capture Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton = New-Object System.Windows.Forms.Button
$OpenCaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenCaptureButton.Location = New-Object System.Drawing.Point(668, 215)
$OpenCaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenCaptureButton.TabIndex = 53
$OpenCaptureButton.Text = "Open CSV Output Folder"
$OpenCaptureButton.UseVisualStyleBackColor = $false
$OpenCaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add( $OpenCaptureButton )
#endregion ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenCaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$OpenCaptureButtonToolTip.AutoPopDelay = 5000
$OpenCaptureButtonToolTip.InitialDelay = 50
$OpenCaptureButtonToolTip.IsBalloon = $true
$OpenCaptureButtonToolTip.ReshowDelay = 100
$OpenCaptureButtonToolTip.SetToolTip( $OpenCaptureButton, "Click once collection is complete to open output folder"+[char]13+[char]10+"seleted above." )
#endregion ~~< OpenCaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add( $TabCapture )
#endregion ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDraw = New-Object System.Windows.Forms.TabPage
$TabDraw.Location = New-Object System.Drawing.Point(4, 22)
$TabDraw.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDraw.Size = New-Object System.Drawing.Size(982, 486)
$TabDraw.TabIndex = 2
$TabDraw.Text = "Draw Visio"
$TabDraw.UseVisualStyleBackColor = $true


#region ~~< Csv Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvInput >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton = New-Object System.Windows.Forms.Button
$DrawCsvInputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCsvInputButton.Location = New-Object System.Drawing.Point(220, 10)
$DrawCsvInputButton.Size = New-Object System.Drawing.Size(750, 25)
$DrawCsvInputButton.TabIndex = 1
$DrawCsvInputButton.Text = "Select CSV Input Folder"
$DrawCsvInputButton.UseVisualStyleBackColor = $false
$DrawCsvInputButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $DrawCsvInputButton )
#endregion ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$DrawCsvInputButtonToolTip.AutoPopDelay = 5000
$DrawCsvInputButtonToolTip.InitialDelay = 50
$DrawCsvInputButtonToolTip.IsBalloon = $true
$DrawCsvInputButtonToolTip.ReshowDelay = 100
$DrawCsvInputButtonToolTip.SetToolTip( $DrawCsvInputButton, "Click to select the folder where the CSV"+[char]39+"s are located."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green." )
#endregion ~~< DrawCsvInputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputLabel = New-Object System.Windows.Forms.Label
$DrawCsvInputLabel.Font = New-Object System.Drawing.Font( "Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$DrawCsvInputLabel.Location = New-Object System.Drawing.Point(10, 10)
$DrawCsvInputLabel.Size = New-Object System.Drawing.Size(190, 25)
$DrawCsvInputLabel.TabIndex = 0
$DrawCsvInputLabel.Text = "CSV Input Folder:"
$TabDraw.Controls.Add( $DrawCsvInputLabel )
#endregion ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$DrawCsvBrowse.Description = "Select a directory"
$DrawCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< CsvInput >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VirtualCenterCsvValidation = New-Object System.Windows.Forms.Label
$VirtualCenterCsvValidation.Location = New-Object System.Drawing.Point(10, 40)
$VirtualCenterCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$VirtualCenterCsvValidation.TabIndex = 2
$VirtualCenterCsvValidation.Text = "vCenter CSV File:"
$TabDraw.Controls.Add( $VirtualCenterCsvValidation )
#endregion ~~< vCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VirtualCenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$VirtualCenterCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 40)
$VirtualCenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VirtualCenterCsvValidationCheck.TabIndex = 3
$VirtualCenterCsvValidationCheck.Text = ""
$VirtualCenterCsvValidationCheck.add_Click( { VirtualCenterCsvValidationCheckClick( $VirtualCenterCsvValidationCheck ) } )
$TabDraw.Controls.Add( $VirtualCenterCsvValidationCheck )
#endregion ~~< VirtualCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VirtualCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComposerServersCsvValidation = New-Object System.Windows.Forms.Label
$ComposerServersCsvValidation.Location = New-Object System.Drawing.Point(10, 60)
$ComposerServersCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$ComposerServersCsvValidation.TabIndex = 4
$ComposerServersCsvValidation.Text = "Composer Servers CSV File:"
$TabDraw.Controls.Add( $ComposerServersCsvValidation )
#endregion ~~< ComposerServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComposerServersCsvValidationCheck = New-Object System.Windows.Forms.Label
$ComposerServersCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 60)
$ComposerServersCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ComposerServersCsvValidationCheck.TabIndex = 5
$ComposerServersCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $ComposerServersCsvValidationCheck )
#endregion ~~< ComposerServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ComposerServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServersCsvValidation = New-Object System.Windows.Forms.Label
$ConnectionServersCsvValidation.Location = New-Object System.Drawing.Point(10, 80)
$ConnectionServersCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$ConnectionServersCsvValidation.TabIndex = 6
$ConnectionServersCsvValidation.Text = "Connection Servers CSV File:"
$TabDraw.Controls.Add( $ConnectionServersCsvValidation )
#endregion ~~< ConnectionServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectionServersCsvValidationCheck = New-Object System.Windows.Forms.Label
$ConnectionServersCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 80)
$ConnectionServersCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ConnectionServersCsvValidationCheck.TabIndex = 7
$ConnectionServersCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $ConnectionServersCsvValidationCheck )
#endregion ~~< ConnectionServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ConnectionServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PoolsCsvValidation = New-Object System.Windows.Forms.Label
$PoolsCsvValidation.Location = New-Object System.Drawing.Point(10, 100)
$PoolsCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$PoolsCsvValidation.TabIndex = 8
$PoolsCsvValidation.Text = "Pools CSV File:"
$TabDraw.Controls.Add( $PoolsCsvValidation )
#endregion ~~< PoolsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PoolsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PoolsCsvValidationCheck = New-Object System.Windows.Forms.Label
$PoolsCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 100)
$PoolsCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$PoolsCsvValidationCheck.TabIndex = 9
$PoolsCsvValidationCheck.Text = ""
$PoolsCsvValidationCheck.add_Click( { PoolsCsvValidationCheckClick( $PoolsCsvValidationCheck ) } )
$TabDraw.Controls.Add( $PoolsCsvValidationCheck )
#endregion ~~< PoolsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PoolsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DesktopsCsvValidation = New-Object System.Windows.Forms.Label
$DesktopsCsvValidation.Location = New-Object System.Drawing.Point(10, 120)
$DesktopsCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$DesktopsCsvValidation.TabIndex = 10
$DesktopsCsvValidation.Text = "Desktops CSV File:"
$TabDraw.Controls.Add( $DesktopsCsvValidation )
#endregion ~~< DesktopsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DesktopsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DesktopsCsvValidationCheck = New-Object System.Windows.Forms.Label
$DesktopsCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 120)
$DesktopsCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DesktopsCsvValidationCheck.TabIndex = 11
$DesktopsCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $DesktopsCsvValidationCheck )
#endregion ~~< DesktopsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< DesktopsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RDSServersCsvValidation = New-Object System.Windows.Forms.Label
$RDSServersCsvValidation.Location = New-Object System.Drawing.Point(10, 140)
$RDSServersCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$RDSServersCsvValidation.TabIndex = 12
$RDSServersCsvValidation.Text = "RDS Servers CSV File:"
$TabDraw.Controls.Add( $RDSServersCsvValidation )
#endregion ~~< RDSServersCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RDSServersCsvValidationCheck = New-Object System.Windows.Forms.Label
$RDSServersCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 140)
$RDSServersCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$RDSServersCsvValidationCheck.TabIndex = 13
$RDSServersCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $RDSServersCsvValidationCheck )
#endregion ~~< RDSServersCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< RDSServersCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FarmsCsvValidation = New-Object System.Windows.Forms.Label
$FarmsCsvValidation.Location = New-Object System.Drawing.Point(10, 160)
$FarmsCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$FarmsCsvValidation.TabIndex = 14
$FarmsCsvValidation.Text = "Farms CSV File:"
$TabDraw.Controls.Add( $FarmsCsvValidation )
#endregion ~~< FarmsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FarmsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FarmsCsvValidationCheck = New-Object System.Windows.Forms.Label
$FarmsCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 160)
$FarmsCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$FarmsCsvValidationCheck.TabIndex = 15
$FarmsCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $FarmsCsvValidationCheck )
#endregion ~~< FarmsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< FarmsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ApplicationsCsvValidation = New-Object System.Windows.Forms.Label
$ApplicationsCsvValidation.Location = New-Object System.Drawing.Point(10, 180)
$ApplicationsCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$ApplicationsCsvValidation.TabIndex = 16
$ApplicationsCsvValidation.Text = "Applications CSV File:"
$TabDraw.Controls.Add( $ApplicationsCsvValidation )
#endregion ~~< ApplicationsCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ApplicationsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ApplicationsCsvValidationCheck = New-Object System.Windows.Forms.Label
$ApplicationsCsvValidationCheck.Location = New-Object System.Drawing.Point(190, 180)
$ApplicationsCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ApplicationsCsvValidationCheck.TabIndex = 17
$ApplicationsCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $ApplicationsCsvValidationCheck )
#endregion ~~< ApplicationsCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ApplicationsCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GatewaysCsvValidation = New-Object System.Windows.Forms.Label
$GatewaysCsvValidation.Location = New-Object System.Drawing.Point(280, 40)
$GatewaysCsvValidation.Size = New-Object System.Drawing.Size(175, 20)
$GatewaysCsvValidation.TabIndex = 18
$GatewaysCsvValidation.Text = "Gateways CSV File:"
$TabDraw.Controls.Add( $GatewaysCsvValidation )
#endregion ~~< GatewaysCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< GatewaysCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$GatewaysCsvValidationCheck = New-Object System.Windows.Forms.Label
$GatewaysCsvValidationCheck.Location = New-Object System.Drawing.Point(460, 40)
$GatewaysCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$GatewaysCsvValidationCheck.TabIndex = 19
$GatewaysCsvValidationCheck.Text = ""
$TabDraw.Controls.Add( $GatewaysCsvValidationCheck )
#endregion ~~< GatewaysCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< GatewaysCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton = New-Object System.Windows.Forms.Button
$CsvValidationButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CsvValidationButton.Location = New-Object System.Drawing.Point(8, 200)
$CsvValidationButton.Size = New-Object System.Drawing.Size(200, 25)
$CsvValidationButton.TabIndex = 50
$CsvValidationButton.Text = "Check for CSVs"
$CsvValidationButton.UseVisualStyleBackColor = $false
$CsvValidationButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $CsvValidationButton )
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$CsvValidationButtonToolTip.IsBalloon = $true
$CsvValidationButtonToolTip.SetToolTip( $CsvValidationButton, "Click to validate that the required CSV files are present."+[char]13+[char]10+"You must validate files prior to drawing Visio." )
#endregion ~~< CsvValidationButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Csv Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Output Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOutputLabel = New-Object System.Windows.Forms.Label
$VisioOutputLabel.Font = New-Object System.Drawing.Font( "Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte](0) ) )
$VisioOutputLabel.Location = New-Object System.Drawing.Point(10, 230)
$VisioOutputLabel.Size = New-Object System.Drawing.Size(215, 25)
$VisioOutputLabel.TabIndex = 51
$VisioOutputLabel.Text = "Visio Output Folder:"
$TabDraw.Controls.Add( $VisioOutputLabel )
#endregion ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton = New-Object System.Windows.Forms.Button
$VisioOpenOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$VisioOpenOutputButton.Location = New-Object System.Drawing.Point(230, 230)
$VisioOpenOutputButton.Size = New-Object System.Drawing.Size(740, 25)
$VisioOpenOutputButton.TabIndex = 52
$VisioOpenOutputButton.Text = "Select Visio Output Folder"
$VisioOpenOutputButton.UseVisualStyleBackColor = $false
$VisioOpenOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $VisioOpenOutputButton )
#endregion ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$VisioOpenOutputButtonToolTip.AutoPopDelay = 5000
$VisioOpenOutputButtonToolTip.InitialDelay = 50
$VisioOpenOutputButtonToolTip.IsBalloon = $true
$VisioOpenOutputButtonToolTip.ReshowDelay = 100
$VisioOpenOutputButtonToolTip.SetToolTip( $VisioOpenOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"Visio Drawings."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green." )
#endregion ~~< VisioOpenOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$VisioBrowse.Description = "Select a directory"
$VisioBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Output Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Infrastructure >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Infrastructure_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Infrastructure_Complete = New-Object System.Windows.Forms.Label
$Infrastructure_Complete.Location = New-Object System.Drawing.Point(315, 260)
$Infrastructure_Complete.Size = New-Object System.Drawing.Size(150, 20)
$Infrastructure_Complete.TabIndex = 64
$Infrastructure_Complete.Text = ""
$TabDraw.Controls.Add( $Infrastructure_Complete )
#endregion ~~< Infrastructure_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Infrastructure_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Infrastructure_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Infrastructure_DrawCheckBox.Checked = $true
$Infrastructure_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Infrastructure_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 260)
$Infrastructure_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Infrastructure_DrawCheckBox.TabIndex = 63
$Infrastructure_DrawCheckBox.Text = "Infrastructure Visio Drawing"
$Infrastructure_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $Infrastructure_DrawCheckBox )
#endregion ~~< Infrastructure_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Infrastructure_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Infrastructure_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$Infrastructure_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Infrastructure_DrawCheckBoxToolTip.InitialDelay = 50
$Infrastructure_DrawCheckBoxToolTip.IsBalloon = $true
$Infrastructure_DrawCheckBoxToolTip.ReshowDelay = 100
$Infrastructure_DrawCheckBoxToolTip.SetToolTip( $Infrastructure_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Infrastructure."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< Infrastructure_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Infrastructure >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Pool_to_Desktop_Complete = New-Object System.Windows.Forms.Label
$Pool_to_Desktop_Complete.Location = New-Object System.Drawing.Point(315, 280)
$Pool_to_Desktop_Complete.Size = New-Object System.Drawing.Size(150, 20)
$Pool_to_Desktop_Complete.TabIndex = 54
$Pool_to_Desktop_Complete.Text = ""
$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
#endregion ~~< Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Pool_to_Desktop_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Pool_to_Desktop_DrawCheckBox.Checked = $true
$Pool_to_Desktop_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Pool_to_Desktop_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 280)
$Pool_to_Desktop_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Pool_to_Desktop_DrawCheckBox.TabIndex = 53
$Pool_to_Desktop_DrawCheckBox.Text = "Pool to Desktop Visio Drawing"
$Pool_to_Desktop_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $Pool_to_Desktop_DrawCheckBox )
#endregion ~~< Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Pool_to_Desktop_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$Pool_to_Desktop_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Pool_to_Desktop_DrawCheckBoxToolTip.InitialDelay = 50
$Pool_to_Desktop_DrawCheckBoxToolTip.IsBalloon = $true
$Pool_to_Desktop_DrawCheckBoxToolTip.ReshowDelay = 100
$Pool_to_Desktop_DrawCheckBoxToolTip.SetToolTip( $Pool_to_Desktop_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Pool to Desktop. This will also add all"+[char]13+[char]10+"metadata to the Visio shapes."+[char]13+[char]10 )
#endregion ~~< Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Desktop_Complete = New-Object System.Windows.Forms.Label
$LC_Pool_to_Desktop_Complete.Location = New-Object System.Drawing.Point(315, 300)
$LC_Pool_to_Desktop_Complete.Size = New-Object System.Drawing.Size(150, 20)
$LC_Pool_to_Desktop_Complete.TabIndex = 56
$LC_Pool_to_Desktop_Complete.Text = ""
$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
#endregion ~~< LC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Desktop_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$LC_Pool_to_Desktop_DrawCheckBox.Checked = $true
$LC_Pool_to_Desktop_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$LC_Pool_to_Desktop_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 300)
$LC_Pool_to_Desktop_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$LC_Pool_to_Desktop_DrawCheckBox.TabIndex = 55
$LC_Pool_to_Desktop_DrawCheckBox.Text = "Linked Clone Pool to Desktop Visio Drawing"
$LC_Pool_to_Desktop_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $LC_Pool_to_Desktop_DrawCheckBox )
#endregion ~~< LC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Desktop_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$LC_Pool_to_Desktop_DrawCheckBoxToolTip.AutoPopDelay = 5000
$LC_Pool_to_Desktop_DrawCheckBoxToolTip.InitialDelay = 50
$LC_Pool_to_Desktop_DrawCheckBoxToolTip.IsBalloon = $true
$LC_Pool_to_Desktop_DrawCheckBoxToolTip.ReshowDelay = 100
$LC_Pool_to_Desktop_DrawCheckBoxToolTip.SetToolTip( $LC_Pool_to_Desktop_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Linked Clone Pool to Desktop. This will also add all "+[char]13+[char]10+"metadata to the Visio shapes."+[char]13+[char]10 )
#endregion ~~< LC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< LC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Application_Complete = New-Object System.Windows.Forms.Label
$LC_Pool_to_Application_Complete.Location = New-Object System.Drawing.Point(315, 320)
$LC_Pool_to_Application_Complete.Size = New-Object System.Drawing.Size(150, 20)
$LC_Pool_to_Application_Complete.TabIndex = 57
$LC_Pool_to_Application_Complete.Text = ""
$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
#endregion ~~< LC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Application_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$LC_Pool_to_Application_DrawCheckBox.Checked = $true
$LC_Pool_to_Application_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$LC_Pool_to_Application_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 320)
$LC_Pool_to_Application_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$LC_Pool_to_Application_DrawCheckBox.TabIndex = 58
$LC_Pool_to_Application_DrawCheckBox.Text = "Linked Clone Pool to Application Visio Drawing"
$LC_Pool_to_Application_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $LC_Pool_to_Application_DrawCheckBox )
#endregion ~~< LC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LC_Pool_to_Application_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$LC_Pool_to_Application_DrawCheckBoxToolTip.AutoPopDelay = 5000
$LC_Pool_to_Application_DrawCheckBoxToolTip.InitialDelay = 50
$LC_Pool_to_Application_DrawCheckBoxToolTip.IsBalloon = $true
$LC_Pool_to_Application_DrawCheckBoxToolTip.ReshowDelay = 100
$LC_Pool_to_Application_DrawCheckBoxToolTip.SetToolTip( $LC_Pool_to_Application_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Linked Clone Pool to Application. This will also add all "+[char]13+[char]10+"metadata to the Visio shapes."+[char]13+[char]10 )
#endregion ~~< LC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< LC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Desktop_Complete = New-Object System.Windows.Forms.Label
$FC_Pool_to_Desktop_Complete.Location = New-Object System.Drawing.Point(315, 340)
$FC_Pool_to_Desktop_Complete.Size = New-Object System.Drawing.Size(150, 20)
$FC_Pool_to_Desktop_Complete.TabIndex = 60
$FC_Pool_to_Desktop_Complete.Text = ""
$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
#endregion ~~< FC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Desktop_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$FC_Pool_to_Desktop_DrawCheckBox.Checked = $true
$FC_Pool_to_Desktop_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FC_Pool_to_Desktop_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 340)
$FC_Pool_to_Desktop_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$FC_Pool_to_Desktop_DrawCheckBox.TabIndex = 59
$FC_Pool_to_Desktop_DrawCheckBox.Text = "Full Clone Pool to Desktop Visio Drawing"
$FC_Pool_to_Desktop_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $FC_Pool_to_Desktop_DrawCheckBox )
#endregion ~~< FC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Desktop_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$FC_Pool_to_Desktop_DrawCheckBoxToolTip.AutoPopDelay = 5000
$FC_Pool_to_Desktop_DrawCheckBoxToolTip.InitialDelay = 50
$FC_Pool_to_Desktop_DrawCheckBoxToolTip.IsBalloon = $true
$FC_Pool_to_Desktop_DrawCheckBoxToolTip.ReshowDelay = 100
$FC_Pool_to_Desktop_DrawCheckBoxToolTip.SetToolTip( $FC_Pool_to_Desktop_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Full Clone Pool to Desktop."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< FC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< FC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Application_Complete = New-Object System.Windows.Forms.Label
$FC_Pool_to_Application_Complete.Location = New-Object System.Drawing.Point(315, 360)
$FC_Pool_to_Application_Complete.Size = New-Object System.Drawing.Size(150, 20)
$FC_Pool_to_Application_Complete.TabIndex = 62
$FC_Pool_to_Application_Complete.Text = ""
$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
#endregion ~~< FC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Application_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$FC_Pool_to_Application_DrawCheckBox.Checked = $true
$FC_Pool_to_Application_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FC_Pool_to_Application_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 360)
$FC_Pool_to_Application_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$FC_Pool_to_Application_DrawCheckBox.TabIndex = 61
$FC_Pool_to_Application_DrawCheckBox.Text = "Full Clone Pool to Application Visio Drawing"
$FC_Pool_to_Application_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $FC_Pool_to_Application_DrawCheckBox )
#endregion ~~< FC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FC_Pool_to_Application_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$FC_Pool_to_Application_DrawCheckBoxToolTip.AutoPopDelay = 5000
$FC_Pool_to_Application_DrawCheckBoxToolTip.InitialDelay = 50
$FC_Pool_to_Application_DrawCheckBoxToolTip.IsBalloon = $true
$FC_Pool_to_Application_DrawCheckBoxToolTip.ReshowDelay = 100
$FC_Pool_to_Application_DrawCheckBoxToolTip.SetToolTip( $FC_Pool_to_Application_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Full Clone Pool to Application."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< FC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< FC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Desktop_Complete = New-Object System.Windows.Forms.Label
$IC_Pool_to_Desktop_Complete.Location = New-Object System.Drawing.Point(315, 380)
$IC_Pool_to_Desktop_Complete.Size = New-Object System.Drawing.Size(150, 20)
$IC_Pool_to_Desktop_Complete.TabIndex = 64
$IC_Pool_to_Desktop_Complete.Text = ""
$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
#endregion ~~< IC_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Desktop_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$IC_Pool_to_Desktop_DrawCheckBox.Checked = $true
$IC_Pool_to_Desktop_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$IC_Pool_to_Desktop_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 380)
$IC_Pool_to_Desktop_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$IC_Pool_to_Desktop_DrawCheckBox.TabIndex = 63
$IC_Pool_to_Desktop_DrawCheckBox.Text = "Instant Clones Pool to Desktop Visio Drawing"
$IC_Pool_to_Desktop_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $IC_Pool_to_Desktop_DrawCheckBox )
#endregion ~~< IC_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Desktop_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$IC_Pool_to_Desktop_DrawCheckBoxToolTip.AutoPopDelay = 5000
$IC_Pool_to_Desktop_DrawCheckBoxToolTip.InitialDelay = 50
$IC_Pool_to_Desktop_DrawCheckBoxToolTip.IsBalloon = $true
$IC_Pool_to_Desktop_DrawCheckBoxToolTip.ReshowDelay = 100
$IC_Pool_to_Desktop_DrawCheckBoxToolTip.SetToolTip( $IC_Pool_to_Desktop_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Instant Clone Pool to Desktop."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< IC_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< IC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Application_Complete = New-Object System.Windows.Forms.Label
$IC_Pool_to_Application_Complete.Location = New-Object System.Drawing.Point(315, 400)
$IC_Pool_to_Application_Complete.Size = New-Object System.Drawing.Size(150, 20)
$IC_Pool_to_Application_Complete.TabIndex = 66
$IC_Pool_to_Application_Complete.Text = ""
$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
#endregion ~~< IC_Pool_to_Application_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Application_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$IC_Pool_to_Application_DrawCheckBox.Checked = $true
$IC_Pool_to_Application_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$IC_Pool_to_Application_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 400)
$IC_Pool_to_Application_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$IC_Pool_to_Application_DrawCheckBox.TabIndex = 65
$IC_Pool_to_Application_DrawCheckBox.Text = "Instant Clones Pool to Application Visio Drawing"
$IC_Pool_to_Application_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $IC_Pool_to_Application_DrawCheckBox )
#endregion ~~< IC_Pool_to_Application_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IC_Pool_to_Application_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$IC_Pool_to_Application_DrawCheckBoxToolTip.AutoPopDelay = 5000
$IC_Pool_to_Application_DrawCheckBoxToolTip.InitialDelay = 50
$IC_Pool_to_Application_DrawCheckBoxToolTip.IsBalloon = $true
$IC_Pool_to_Application_DrawCheckBoxToolTip.ReshowDelay = 100
$IC_Pool_to_Application_DrawCheckBoxToolTip.SetToolTip( $IC_Pool_to_Application_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Instant Clone Pool to Application."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< IC_Pool_to_Application_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< IC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Unmanaged_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Unmanaged_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Unmanaged_Pool_to_Desktop_Complete = New-Object System.Windows.Forms.Label
$Unmanaged_Pool_to_Desktop_Complete.Location = New-Object System.Drawing.Point(315, 420)
$Unmanaged_Pool_to_Desktop_Complete.Size = New-Object System.Drawing.Size(150, 20)
$Unmanaged_Pool_to_Desktop_Complete.TabIndex = 68
$Unmanaged_Pool_to_Desktop_Complete.Text = ""
$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
#endregion ~~< Unmanaged_Pool_to_Desktop_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Unmanaged_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Unmanaged_Pool_to_Desktop_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Unmanaged_Pool_to_Desktop_DrawCheckBox.Checked = $true
$Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Unmanaged_Pool_to_Desktop_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 420)
$Unmanaged_Pool_to_Desktop_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Unmanaged_Pool_to_Desktop_DrawCheckBox.TabIndex = 67
$Unmanaged_Pool_to_Desktop_DrawCheckBox.Text = "Unmanaged Pool to Desktop Visio Drawing"
$Unmanaged_Pool_to_Desktop_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_DrawCheckBox )
#endregion ~~< Unmanaged_Pool_to_Desktop_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip.InitialDelay = 50
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip.IsBalloon = $true
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip.ReshowDelay = 100
$Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip.SetToolTip( $Unmanaged_Pool_to_Desktop_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Unmanaged Pool to Desktop."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< Unmanaged_Pool_to_Desktop_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Unmanaged_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farm_to_Remote_Desktop_Services >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farm_to_Remote_Desktop_Services_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Farm_to_Remote_Desktop_Services_Complete = New-Object System.Windows.Forms.Label
$Farm_to_Remote_Desktop_Services_Complete.Location = New-Object System.Drawing.Point(840, 260)
$Farm_to_Remote_Desktop_Services_Complete.Size = New-Object System.Drawing.Size(150, 20)
$Farm_to_Remote_Desktop_Services_Complete.TabIndex = 70
$Farm_to_Remote_Desktop_Services_Complete.Text = ""
$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
#endregion ~~< Farm_to_Remote_Desktop_Services_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farm_to_Remote_Desktop_Services_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Farm_to_Remote_Desktop_Services_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Farm_to_Remote_Desktop_Services_DrawCheckBox.Checked = $true
$Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Farm_to_Remote_Desktop_Services_DrawCheckBox.Location = New-Object System.Drawing.Point(485, 260)
$Farm_to_Remote_Desktop_Services_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$Farm_to_Remote_Desktop_Services_DrawCheckBox.TabIndex = 69
$Farm_to_Remote_Desktop_Services_DrawCheckBox.Text = "Farm to Remote Desktop Services Visio Drawing"
$Farm_to_Remote_Desktop_Services_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_DrawCheckBox )
#endregion ~~< Farm_to_Remote_Desktop_Services_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip.InitialDelay = 50
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip.IsBalloon = $true
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip.ReshowDelay = 100
$Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip.SetToolTip( $Farm_to_Remote_Desktop_Services_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"Farm to Remote Desktop Services."+[char]13+[char]10+"This will also add all metadata to the Visio shapes." )
#endregion ~~< Farm_to_Remote_Desktop_Services_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Farm_to_Remote_Desktop_Services >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton = New-Object System.Windows.Forms.Button
$DrawUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawUncheckButton.Location = New-Object System.Drawing.Point(8, 450)
$DrawUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawUncheckButton.TabIndex = 87
$DrawUncheckButton.Text = "Uncheck All"
$DrawUncheckButton.UseVisualStyleBackColor = $false
$DrawUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $DrawUncheckButton )
#endregion ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$DrawUncheckButtonToolTip.AutoPopDelay = 5000
$DrawUncheckButtonToolTip.InitialDelay = 50
$DrawUncheckButtonToolTip.IsBalloon = $true
$DrawUncheckButtonToolTip.ReshowDelay = 100
$DrawUncheckButtonToolTip.SetToolTip( $DrawUncheckButton, "Click to clear all check boxes above." )
#endregion ~~< DrawUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton = New-Object System.Windows.Forms.Button
$DrawCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCheckButton.Location = New-Object System.Drawing.Point(228, 450)
$DrawCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawCheckButton.TabIndex = 88
$DrawCheckButton.Text = "Check All"
$DrawCheckButton.UseVisualStyleBackColor = $false
$DrawCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $DrawCheckButton )
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$DrawCheckButtonToolTip.AutoPopDelay = 5000
$DrawCheckButtonToolTip.InitialDelay = 50
$DrawCheckButtonToolTip.IsBalloon = $true
$DrawCheckButtonToolTip.ReshowDelay = 100
$DrawCheckButtonToolTip.SetToolTip( $DrawCheckButton, "Click to check all check boxes above." )
#endregion ~~< DrawCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton = New-Object System.Windows.Forms.Button
$DrawButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawButton.Location = New-Object System.Drawing.Point(448, 450)
$DrawButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawButton.TabIndex = 89
$DrawButton.Text = "Draw Visio"
$DrawButton.UseVisualStyleBackColor = $false
$DrawButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $DrawButton )
#endregion ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$DrawButtonToolTip.AutoPopDelay = 5000
$DrawButtonToolTip.InitialDelay = 50
$DrawButtonToolTip.IsBalloon = $true
$DrawButtonToolTip.ReshowDelay = 100
$DrawButtonToolTip.SetToolTip( $DrawButton, "Click to begin drawing environment based on"+[char]13+[char]10+"options selected above." )
#endregion ~~< DrawButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Draw Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open Visio Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton = New-Object System.Windows.Forms.Button
$OpenVisioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenVisioButton.Location = New-Object System.Drawing.Point(668, 450)
$OpenVisioButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenVisioButton.TabIndex = 90
$OpenVisioButton.Text = "Open Visio Drawing"
$OpenVisioButton.UseVisualStyleBackColor = $false
$OpenVisioButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add( $OpenVisioButton )
#endregion ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenVisioButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButtonToolTip = New-Object System.Windows.Forms.ToolTip( $components )
$OpenVisioButtonToolTip.AutoPopDelay = 5000
$OpenVisioButtonToolTip.InitialDelay = 50
$OpenVisioButtonToolTip.IsBalloon = $true
$OpenVisioButtonToolTip.ReshowDelay = 100
$OpenVisioButtonToolTip.SetToolTip( $OpenVisioButton, "Click to open Visio drawing once all above check boxes"+[char]13+[char]10+"are marked as completed." )
#endregion ~~< OpenVisioButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open Visio Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add( $TabDraw)
#endregion ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.SelectedIndex = 0
$vDiagram.Controls.Add( $LowerTabs)

#endregion ~~< LowerTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Form Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellCheck = $PSVersionTable.PSVersion
if ( $PowershellCheck.Major -ge 4 ) `
{ `
	$PowershellInstalled.Forecolor = "Green"
	$PowershellInstalled.Text = "Installed Version $PowershellCheck"
}
else `
{ `
	$PowershellInstalled.Forecolor = "Red"
	$PowershellInstalled.Text = "Not installed or Powershell version lower than 4"
}
#endregion ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleCheck = ( Get-Module VMware.PowerCLI -ListAvailable | Where-Object { $_.Name -eq "VMware.PowerCLI" } | Sort-Object Version -Descending )
$PowerCliModuleVersion = ( $PowerCliModuleCheck.Version[0] )
if ( $null -ne $PowerCliModuleCheck ) `
{ `
	$PowerCliModuleInstalled.Forecolor = "Green"
	$PowerCliModuleInstalled.Text = "Installed Version $PowerCliModuleVersion"
}
else `
{ `
	$PowerCliModuleInstalled.Forecolor = "Red"
	$PowerCliModuleInstalled.Text = "Not Installed"
}
#endregion ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( $null -ne ( Get-PSSnapin -registered | Where-Object { $_.Name -eq "VMware.VimAutomation.Core" } ) ) `
{ `
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerClI Installed"
}
elseif ( $null -ne $PowerCliModuleCheck ) `
{ `
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerCLI Module Installed"
}
else `
{ `
	$PowerCliInstalled.Forecolor = "Red"
	$PowerCliInstalled.Text = "PowerCLI or PowerCli Module not installed"
}
#endregion ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( ( Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) -or $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) ) `
{ `
	$VisioInstalled.Forecolor = "Green"
	$VisioInstalled.Text = "Installed"
}
else `
{ `
	$VisioInstalled.Forecolor = "Red"
	$VisioInstalled.Text = "Visio is Not Installed"
	[System.Windows.Forms.MessageBox]::Show("Visio is not installed on this machine.
	
In order to complete the draw process Visio is required, you can capture from this machine and export the outputs to a machine with Visio installed then re-run and point the tool to the folder to continue the draw process.
	
When drawing on another machine please make sure to enter the Connection Server name exactly as it was entered during the capture phase.", "Warning Visio Not Found",0,48)
	
}
#endregion ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton.Add_MouseClick( `
{ `
	$Connected = $global:DefaultHVServers.ExtensionData.Session.Client.ServiceUri ; 
	if ( $Connected -eq $null ) `
	{ `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Red ; 
		$ConnectButton.Text = "Unable to Connect"
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Unable to connect to Horizon" -ForegroundColor Red
		}
		if ( $logcapture -eq $true ) `
		{ `
			$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
			$LogCapturePath = $FileDateTime + " " + $ConnServ + " - vDiagram_Capture.log"
			Start-Transcript -Path "$LogCapturePath"
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Powershell Module version installed:" $PSVersionTable.PSVersion -ForegroundColor Green
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] PowerCLI Module versions installed:" $PowerCliModuleCheck.Version -ForegroundColor Green
		}
	}
	else `
	{ `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green ;
		$ConnectButton.Text = "Connected to $global:DefaultHVServers."
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Connected to $global:DefaultHVServers." -ForegroundColor Green
		}
		if ( $logcapture -eq $true ) `
		{ `
			$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
			$LogCapturePath = $FileDateTime + " " + $ConnServ + " - vDiagram_Capture.log"
			Start-Transcript -Path "$LogCapturePath"
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Powershell Module version installed:" $PSVersionTable.PSVersion -ForegroundColor Green
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] PowerCLI Module versions installed:" $PowerCliModuleCheck.Version -ForegroundColor Green
		}
	}
} )
$ConnectButton.Add_Click( { Connect_HVServer } )
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton.Add_Click( `
{ `
	Find_CaptureCsvFolder ; 
	if ( $CaptureCsvFolder -eq $null ) `
	{ `
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$CaptureCsvOutputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$CaptureCsvOutputButton.Text = $CaptureCsvFolder
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected CSV export folder = $CaptureCsvFolder" -ForegroundColor Green
		}
	}
	Check_CaptureCsvFolder
} )
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton.Add_Click( `
{ `
	$VirtualCenterCsvCheckBox.CheckState = "UnChecked" ;
	$ComposerServersCsvCheckBox.CheckState = "UnChecked" ;
	$ConnectionServersCsvCheckBox.CheckState = "UnChecked" ;
	$PoolsCsvCheckBox.CheckState = "UnChecked" ;
	$DesktopsCsvCheckBox.CheckState = "UnChecked" ;
	$RDSServersCsvCheckBox.CheckState = "UnChecked" ;
	$FarmsCsvCheckBox.CheckState = "UnChecked" ;
	$ApplicationsCsvCheckBox.CheckState = "UnChecked" ;
	$GatewaysCsvCheckBox.CheckState = "UnChecked" ;
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Capture CSV Uncheck All selected." -ForegroundColor Magenta
	}	
} )
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton.Add_Click( `
{ `
	$VirtualCenterCsvCheckBox.CheckState = "Checked" ;
	$ComposerServersCsvCheckBox.CheckState = "Checked" ;
	$ConnectionServersCsvCheckBox.CheckState = "Checked" ;
	$PoolsCsvCheckBox.CheckState = "Checked" ;
	$DesktopsCsvCheckBox.CheckState = "Checked" ;
	$RDSServersCsvCheckBox.CheckState = "Checked" ;
	$FarmsCsvCheckBox.CheckState = "Checked" ;
	$ApplicationsCsvCheckBox.CheckState = "Checked" ;
	$GatewaysCsvCheckBox.CheckState = "Checked" ;
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Capture CSV Check All selected." -ForegroundColor Magenta
	}
} )
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton.Add_Click( `
{ `
	if( $CaptureCsvFolder -eq $null ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
		$CaptureButton.Forecolor = [System.Drawing.Color]::Red; 
		$CaptureButton.Text = "Folder Not Selected"
	}
	else `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] CSV collection started." -ForegroundColor Magenta
		}
		if ( $VirtualCenterCsvCheckBox.Checked -eq "True" ) `
		{ `
			$VirtualCenterCsvValidationComplete.Forecolor = "Blue"
			$VirtualCenterCsvValidationComplete.Text = "Processing ....."
			VirtualCenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$VirtualCenterExportFileComplete = $CsvCompleteDir + "-VirtualCenterExport.csv"
			$VirtualCenterCsvComplete = Test-Path $VirtualCenterExportFileComplete
			
			if ( $VirtualCenterCsvComplete -eq $True ) `
			{ `
				$VirtualCenterCsvValidationComplete.Forecolor = "Green"
				$VirtualCenterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VirtualCenterCsvValidationComplete.Forecolor = "Red"
				$VirtualCenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		Connect_HVServer
		$Connected = $global:DefaultHVServers
		
		if ( $Connected -eq $null ) { Connect_HVServer } `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green
		$ConnectButton.Text = "Connected to $global:DefaultHVServers"
		
		if ( $ComposerServersCsvCheckBox.Checked -eq "True" ) `
		{ `
			ComposerServers_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$ComposerServersExportFileComplete = $CsvCompleteDir + "-ComposerServersExport.csv"
			$ComposerServersCsvComplete = Test-Path $ComposerServersExportFileComplete
			
			if ( $ComposerServersCsvComplete -eq $True ) `
			{ `
				$ComposerServersCsvValidationComplete.Forecolor = "Green"
				$ComposerServersCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$ComposerServersCsvValidationComplete.Forecolor = "Red"
				$ComposerServersCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $ConnectionServersCsvCheckBox.Checked -eq "True" ) `
		{ `
			ConnectionServers_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$ConnectionServersExportFileComplete = $CsvCompleteDir + "-ConnectionServersExport.csv"
			$ConnectionServersCsvComplete = Test-Path $ConnectionServersExportFileComplete
			
			if ( $ConnectionServersCsvComplete -eq $True ) `
			{ `
				$ConnectionServersCsvValidationComplete.Forecolor = "Green"
				$ConnectionServersCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$ConnectionServersCsvValidationComplete.Forecolor = "Red"
				$ConnectionServersCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $PoolsCsvCheckBox.Checked -eq "True" ) `
		{ `
			Pools_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$PoolsExportFileComplete = $CsvCompleteDir + "-PoolsExport.csv"
			$PoolsCsvComplete = Test-Path $PoolsExportFileComplete
			
			if ( $PoolsCsvComplete -eq $True ) `
			{ `
				$PoolsCsvValidationComplete.Forecolor = "Green"
				$PoolsCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$PoolsCsvValidationComplete.Forecolor = "Red"
				$PoolsCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DesktopsCsvCheckBox.Checked -eq "True" ) `
		{ `
			Desktops_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$DesktopsExportFileComplete = $CsvCompleteDir + "-DesktopsExport.csv"
			$DesktopsCsvComplete = Test-Path $DesktopsExportFileComplete
			
			if ( $DesktopsCsvComplete -eq $True ) `
			{ `
				$DesktopsCsvValidationComplete.Forecolor = "Green"
				$DesktopsCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DesktopsCsvValidationComplete.Forecolor = "Red"
				$DesktopsCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $RDSServersCsvCheckBox.Checked -eq "True" ) `
		{ `
			RDSServers_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$RDSServersExportFileComplete = $CsvCompleteDir + "-RDSServersExport.csv"
			$RDSServersCsvComplete = Test-Path $RDSServersExportFileComplete
			
			if ( $RDSServersCsvComplete -eq $True ) `
			{ `
				$RDSServersCsvValidationComplete.Forecolor = "Green"
				$RDSServersCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$RDSServersCsvValidationComplete.Forecolor = "Red"
				$RDSServersCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $FarmsCsvCheckBox.Checked -eq "True" ) `
		{ `
			Farms_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$FarmsExportFileComplete = $CsvCompleteDir + "-FarmsExport.csv"
			$FarmsCsvComplete = Test-Path $FarmsExportFileComplete
			
			if ( $FarmsCsvComplete -eq $True ) `
			{ `
				$FarmsCsvValidationComplete.Forecolor = "Green"
				$FarmsCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$FarmsCsvValidationComplete.Forecolor = "Red"
				$FarmsCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $ApplicationsCsvCheckBox.Checked -eq "True" ) `
		{ `
			Applications_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$ApplicationsExportFileComplete = $CsvCompleteDir + "-ApplicationsExport.csv"
			$ApplicationsCsvComplete = Test-Path $ApplicationsExportFileComplete
			
			if ( $ApplicationsCsvComplete -eq $True ) `
			{ `
				$ApplicationsCsvValidationComplete.Forecolor = "Green"
				$ApplicationsCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$ApplicationsCsvValidationComplete.Forecolor = "Red"
				$ApplicationsCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $GatewaysCsvCheckBox.Checked -eq "True" ) `
		{ `
			Gateways_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
			$GatewaysExportFileComplete = $CsvCompleteDir + "-GatewaysExport.csv"
			$GatewaysCsvComplete = Test-Path $GatewaysExportFileComplete
			
			if ( $GatewaysCsvComplete -eq $True ) `
			{ `
				$GatewaysCsvValidationComplete.Forecolor = "Green"
				$GatewaysCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$GatewaysCsvValidationComplete.Forecolor = "Red"
				$GatewaysCsvValidationComplete.Text = "Not Complete"
			}
		}
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss";`
		Write-Host "[$DateTime] CSV capture complete. Please proceed to drawing Visio." -ForegroundColor Green
	
		Disconnect_HVServer
		if ( $logcapture -eq $true ) `
		{ `
			Stop-Transcript
		}
		$ConnectButton.Forecolor = [System.Drawing.Color]::Red
		$ConnectButton.Text = "Disconnected"
		$CaptureButton.Forecolor = [System.Drawing.Color]::Green ; $CaptureButton.Text = "CSV Collection Complete"
	}
} )
#endregion ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton.Add_Click( `
{ `
	Open_Capture_Folder;
	$ConnServTextBox.Text = "" ;
	$UserNameTextBox.Text = "" ;
	$PasswordTextBox.Text = "" ;
	$PasswordTextBox.UseSystemPasswordChar = $true ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to Horizon" ;
	$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Black ;
	$CaptureCsvOutputButton.Text = "Select Output Folder" ;
	$CaptureButton.Forecolor = [System.Drawing.Color]::Black ;
	$CaptureButton.Text = "Collect CSV Data" ;
	$VirtualCenterCsvCheckBox.CheckState = "Checked" ;
	$VirtualCenterCsvValidationComplete.Text = "" ;
	$ComposerServersCsvCheckBox.CheckState = "Checked" ;
	$ComposerServersCsvValidationComplete.Text = "" ;
	$ConnectionServersCsvCheckBox.CheckState = "Checked" ;
	$ConnectionServersCsvValidationComplete.Text = "" ;
	$PoolsCsvCheckBox.CheckState = "Checked" ;
	$PoolsCsvValidationComplete.Text = "" ;
	$DesktopsCsvCheckBox.CheckState = "Checked" ;
	$DesktopsCsvValidationComplete.Text = "" ;
	$RDSServersCsvCheckBox.CheckState = "Checked" ;
	$RDSServersCsvValidationComplete.Text = "" ;
	$FarmsCsvCheckBox.CheckState = "Checked" ;
	$FarmsCsvValidationComplete.Text = "" ;
	$ApplicationsCsvCheckBox.CheckState = "Checked" ;
	$ApplicationsCsvValidationComplete.Text = "" ;
	$GatewaysCsvCheckBox.CheckState = "Checked" ;
	$GatewaysCsvValidationComplete.Text = "" ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to Horizon"
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Opening CSV folder." -ForegroundColor Magenta
	}
} )
#endregion ~~< CaptureOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton.Add_MouseClick( `
{ `
	Find_DrawCsvFolder ;
	if ( $DrawCsvFolder -eq $null ) `
	{ `
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Red ;
		$DrawCsvInputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
		if ( $logdraw -eq $true ) `
		{ `
			$FileDateTime = ( Get-Date -format "yyyy_MM_dd-HH_mm" )
			$LogDrawPath = $FileDateTime + " " + $ConnServ + " - vDiagram_Draw.log"
			Start-Transcript -Path "$LogDrawPath"
		}
	}
	else `
	{ `
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Green ;
		$DrawCsvInputButton.Text = $DrawCsvFolder
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected import folder = $DrawCsvFolder" -ForegroundColor Magenta
		}
		if ( $logdraw -eq $true ) `
		{ `
			$FileDateTime = ( Get-Date -format "yyyy_MM_dd-HH_mm" )
			$LogDrawPath = $FileDateTime + " " + $ConnServ + " - vDiagram_Draw.log"
			Start-Transcript -Path "$LogDrawPath"
		}
	}
} )
$TabDraw.Controls.Add( $DrawCsvInputButton )
#endregion ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton.Add_Click( `
{ `
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Validating CSVs." -ForegroundColor Magenta
	}
	$CsvInputDir = $DrawCsvFolder+"\"+$ConnServTextBox.Text
	$VirtualCenterExportFile = $CsvInputDir + "-VirtualCenterExport.csv"
	$VirtualCenterCsvExists = Test-Path $VirtualCenterExportFile
	$TabDraw.Controls.Add( $VirtualCenterCsvValidationCheck )
	if ( $VirtualCenterCsvExists -eq $True ) `
	{ `
							
		$VirtualCenterCsvValidationCheck.Forecolor = "Green"
		$VirtualCenterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$VirtualCenterCsvValidationCheck.Forecolor = "Red"
		$VirtualCenterCsvValidationCheck.Text = "Not Present"
	}
	
	$ComposerServersExportFile = $CsvInputDir + "-ComposerServersExport.csv"
	$ComposerServersCsvExists = Test-Path $ComposerServersExportFile
	$TabDraw.Controls.Add( $ComposerServersCsvValidationCheck )
			
	if ( $ComposerServersCsvExists -eq $True ) `
	{ `
		$ComposerServersCsvValidationCheck.Forecolor = "Green"
		$ComposerServersCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$ComposerServersCsvValidationCheck.Forecolor = "Red"
		$ComposerServersCsvValidationCheck.Text = "Not Present"
	}
	
	$ConnectionServersExportFile = $CsvInputDir + "-ConnectionServersExport.csv"
	$ConnectionServersCsvExists = Test-Path $ConnectionServersExportFile
	$TabDraw.Controls.Add( $ConnectionServersCsvValidationCheck )
			
	if ( $ConnectionServersCsvExists -eq $True ) `
	{ `
		$ConnectionServersCsvValidationCheck.Forecolor = "Green"
		$ConnectionServersCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$ConnectionServersCsvValidationCheck.Forecolor = "Red"
		$ConnectionServersCsvValidationCheck.Text = "Not Present"
	}
			
	$PoolsExportFile = $CsvInputDir + "-PoolsExport.csv"
	$PoolsCsvExists = Test-Path $PoolsExportFile
	$TabDraw.Controls.Add( $PoolsCsvValidationCheck )
			
	if ( $PoolsCsvExists -eq $True ) `
	{ `
		$PoolsCsvValidationCheck.Forecolor = "Green"
		$PoolsCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$PoolsCsvValidationCheck.Forecolor = "Red"
		$PoolsCsvValidationCheck.Text = "Not Present"
	}
			
	$DesktopsExportFile = $CsvInputDir + "-DesktopsExport.csv"
	$DesktopsCsvExists = Test-Path $DesktopsExportFile
	$TabDraw.Controls.Add( $DesktopsCsvValidationCheck )
			
	if ( $DesktopsCsvExists -eq $True ) `
	{ `
		$DesktopsCsvValidationCheck.Forecolor = "Green"
        $DesktopsCsvValidationCheck.Text = "Present"
        $DesktopCount = Import-CSV $DesktopsExportFile
        if ( ( $DesktopCount | Where-Object { $_.Source -eq "VIEW_COMPOSER" } ).Count -eq 0 ) `
        { `
            $LC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
        }
        if ( ( $DesktopCount | Where-Object { $_.Source -eq "VIRTUAL_CENTER" } ).Count -eq 0 ) `
        { `
            $FC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
        }
        if ( ( $DesktopCount | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" } ).Count -eq 0 ) `
        { `
            $IC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked" 
        }
        if ( ( $DesktopCount | Where-Object { $_.Source -eq "UNMANAGED" } ).Count -eq 0 ) `
        { `
            $Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked" 
        }
	}
	else `
	{ `
		$DesktopsCsvValidationCheck.Forecolor = "Red"
        $DesktopsCsvValidationCheck.Text = "Not Present"
        $LC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
        $FC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
        $IC_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
        $Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = "Unchecked"
	}
	
	$RDSServersExportFile = $CsvInputDir + "-RDSServersExport.csv"
	$RDSServersCsvExists = Test-Path $RDSServersExportFile
	$TabDraw.Controls.Add( $RDSServersCsvValidationCheck )
			
	if ( $RDSServersCsvExists -eq $True ) `
	{ `
		$RDSServersCsvValidationCheck.Forecolor = "Green"
		$RDSServersCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$RDSServersCsvValidationCheck.Forecolor = "Red"
        $RDSServersCsvValidationCheck.Text = "Not Present"
        $Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = "Unchecked"
	}
			
	$FarmsExportFile = $CsvInputDir + "-FarmsExport.csv"
	$FarmsCsvExists = Test-Path $FarmsExportFile
    $TabDraw.Controls.Add( $FarmsCsvValidationCheck )
			
	if ( $FarmsCsvExists -eq $True ) `
	{ `
		$FarmsCsvValidationCheck.Forecolor = "Green"
		$FarmsCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$FarmsCsvValidationCheck.Forecolor = "Red"
        $FarmsCsvValidationCheck.Text = "Not Present"
        $Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = "Unchecked"
	}
			
	$ApplicationsExportFile = $CsvInputDir + "-ApplicationsExport.csv"
	$ApplicationsCsvExists = Test-Path $ApplicationsExportFile
    $TabDraw.Controls.Add( $ApplicationsCsvValidationCheck )

    if ( $ApplicationsCsvExists -eq $True ) `
	{ `
        $ApplicationsCsvValidationCheck.Forecolor = "Green"
        $ApplicationsCsvValidationCheck.Text = "Present"
        
        $LC_Count = 0
        foreach ( $LC in ( Import-CSV $PoolsExportFile | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.SupportedSessionType -like "*APPLICATION" -and $_.ApplicationCount -gt 0 } ) ) `
        { `
            $LC.Name
            $LC_Count++
        }
        if ( $LC_Count -eq 0 ) `
		{ `
			$LC_Pool_to_Application_DrawCheckBox.CheckState = "Unchecked"
		}

        $FC_Count = 0
        foreach ( $FC in ( Import-CSV $PoolsExportFile | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.SupportedSessionType -like "*APPLICATION" -and $_.ApplicationCount -gt 0 } ) ) `
        { `
            $FC.Name
            $FC_Count++
        }
        if ( $FC_Count -eq 0 ) `
		{ `
			$FC_Pool_to_Application_DrawCheckBox.CheckState = "Unchecked"
		}
	
        $IC_Count = 0
        foreach ( $IC in ( Import-CSV $PoolsExportFile | Where-Object { $_.Source -like "INSTANT_CLONE_ENGINE" -and $_.SupportedSessionType -like "*APPLICATION" -and $_.ApplicationCount -gt 0 } ) ) `
        { `
            $IC_Count++
        }
        if ( $IC_Count -eq 0 ) `
        { `
            $IC_Pool_to_Application_DrawCheckBox.CheckState = "Unchecked"
        } `
	}
	else `
	{ `
		$ApplicationsCsvValidationCheck.Forecolor = "Red"
        $ApplicationsCsvValidationCheck.Text = "Not Present"
	}
			
	$GatewaysExportFile = $CsvInputDir + "-GatewaysExport.csv"
	$GatewaysCsvExists = Test-Path $GatewaysExportFile
    $TabDraw.Controls.Add( $GatewaysCsvValidationCheck )
			
	if ( $GatewaysCsvExists -eq $True ) `
	{ `
		$GatewaysCsvValidationCheck.Forecolor = "Green"
		$GatewaysCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		$GatewaysCsvValidationCheck.Forecolor = "Red"
		$GatewaysCsvValidationCheck.Text = "Not Present"
	}
} )
$CsvValidationButton.Add_MouseClick( { $CsvValidationButton.Forecolor = [System.Drawing.Color]::Green ; $CsvValidationButton.Text = "CSV Validation Complete" } )
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton.Add_MouseClick( `
{ `
	Find_DrawVisioFolder; 
	if( $VisioFolder -eq $null ) `
	{ `
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$VisioOpenOutputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$VisioOpenOutputButton.Text = $VisioFolder
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected Visio export folder = $VisioFolder" -ForegroundColor Magenta
		}
	}
} )
#endregion ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton.Add_Click( `
{ `
	$Infrastructure_DrawCheckBox.CheckState = "UnChecked" ;
	$Pool_to_Desktop_DrawCheckBox.CheckState = "UnChecked" ;
	$LC_Pool_to_Desktop_DrawCheckBox.CheckState = "UnChecked" ;
	$LC_Pool_to_Application_DrawCheckBox.CheckState = "UnChecked" ;
	$FC_Pool_to_Desktop_DrawCheckBox.CheckState = "UnChecked" ;
	$FC_Pool_to_Application_DrawCheckBox.CheckState = "UnChecked" ;
	$IC_Pool_to_Desktop_DrawCheckBox.CheckState = "UnChecked" ;
	$IC_Pool_to_Application_DrawCheckBox.CheckState = "UnChecked" ;
	$Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = "UnChecked"
	$Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = "UnChecked"
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Draw - Uncheck All selected." -ForegroundColor Magenta
	}
} )
#endregion ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton.Add_Click( `
{ `
	$Infrastructure_DrawCheckBox.CheckState = "Checked" ;
	$Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$LC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$LC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$FC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$FC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$IC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$IC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked"
	$Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = "Checked"
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Draw - Check All selected." -ForegroundColor Magenta
	}
} )
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton.Add_Click( `
{ `
	if( $VisioFolder -eq $null ) `
	{ `
		$DrawButton.Forecolor = [System.Drawing.Color]::Red ;
		$DrawButton.Text = "Folder Not Selected"
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Starting drawings." -ForegroundColor Magenta
		}
		$DrawButton.Forecolor = [System.Drawing.Color]::Blue ;
		$DrawButton.Text = "Drawing Please Wait" ;
		Create_Visio_Base;
		if ( $Infrastructure_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$Infrastructure_Complete.Forecolor = "Blue"
			$Infrastructure_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $Infrastructure_Complete )
			Infrastructure
			$Infrastructure_Complete.Forecolor = "Green"
			$Infrastructure_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $Infrastructure_Complete )
		};
		if ( $Pool_to_Desktop_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$Pool_to_Desktop_Complete.Forecolor = "Blue"
			$Pool_to_Desktop_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
			Pool_to_Desktop
			$Pool_to_Desktop_Complete.Forecolor = "Green"
			$Pool_to_Desktop_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )

		};
		if ( $LC_Pool_to_Desktop_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$LC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$LC_Pool_to_Desktop_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
			LC_Pool_to_Desktop
			$LC_Pool_to_Desktop_Complete.Forecolor = "Green"
			$LC_Pool_to_Desktop_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
		};
		if ( $LC_Pool_to_Application_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$LC_Pool_to_Application_Complete.Forecolor = "Blue"
			$LC_Pool_to_Application_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
			LC_Pool_to_Application
			$LC_Pool_to_Application_Complete.Forecolor = "Green"
			$LC_Pool_to_Application_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
		};
		if ( $FC_Pool_to_Desktop_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$FC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$FC_Pool_to_Desktop_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
			FC_Pool_to_Desktop
			$FC_Pool_to_Desktop_Complete.Forecolor = "Green"
			$FC_Pool_to_Desktop_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
		};
		if ( $FC_Pool_to_Application_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$FC_Pool_to_Application_Complete.Forecolor = "Blue"
			$FC_Pool_to_Application_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
			FC_Pool_to_Application
			$FC_Pool_to_Application_Complete.Forecolor = "Green"
			$FC_Pool_to_Application_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
		};
		if ( $IC_Pool_to_Desktop_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$IC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$IC_Pool_to_Desktop_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
			IC_Pool_to_Desktop
			$IC_Pool_to_Desktop_Complete.Forecolor = "Green"
			$IC_Pool_to_Desktop_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
		};
		if ( $IC_Pool_to_Application_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$IC_Pool_to_Application_Complete.Forecolor = "Blue"
			$IC_Pool_to_Application_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
			IC_Pool_to_Application
			$IC_Pool_to_Application_Complete.Forecolor = "Green"
			$IC_Pool_to_Application_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
		};
		if ( $Unmanaged_Pool_to_Desktop_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$Unmanaged_Pool_to_Desktop_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
			Unmanaged_Pool_to_Desktop
			$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Green"
			$Unmanaged_Pool_to_Desktop_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
		};
		if ( $Farm_to_Remote_Desktop_Services_DrawCheckBox.Checked -eq "True" ) `
		{ `
			$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
			$Farm_to_Remote_Desktop_Services_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
			Farm_to_Remote_Desktop_Services
			$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Green"
			$Farm_to_Remote_Desktop_Services_Complete.Text = "Complete"
			$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		};
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss" ; `
	Write-Host "[$DateTime] Visio Drawings Complete. Click Open Visio Drawing button to proceed." -ForegroundColor Yellow ; `
	$DrawButton.Forecolor = [System.Drawing.Color]::Green; $DrawButton.Text = "Visio Drawings Complete" `
	} `
	
	# Follow us on Twitter Prompt
	$LikeUs =  [System.Windows.Forms.MessageBox]::Show( "Did you find this script helpful? Click 'Yes' to follow us on Twitter and 'No' cancel.","Follow us on Twitter.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning )
	switch  ( $LikeUs ) `
		{ `
			'Yes' 
			{ `
				Start-Process 'https://twitter.com/vDiagramProject'
				[System.Windows.Forms.MessageBox]::Show( "Your Visio Drawing is now complete. Please validate all drawings to ensure items are not missing.

Please click on the Open Visio Drawing button now to open and compress the file.","vDiagram is now complete!",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information )
			}
			'No'
			{ `
				[System.Windows.Forms.MessageBox]::Show( "Your Visio Drawing is now complete. Please validate all drawings to ensure items are not missing.

Please click on the Open Visio Drawing button now to open and compress the file.","vDiagram is now complete!",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information )
			}
		}

	if ( $logdraw -eq $true ) `
	{ `
		Stop-Transcript
	}
} )
#endregion ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton.Add_Click( `
{ `
	Open_Final_Visio ;
	$ConnServTextBox.Text = "" ;
	$UserNameTextBox.Text = "" ;
	$PasswordTextBox.Text = "" ;
	$PasswordTextBox.UseSystemPasswordChar = $true ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to Horizon" ;
	$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Black ;
	$DrawCsvInputButton.Text = "Select CSV Input Folder" ;
	$VirtualCenterCsvValidationCheck.Text = "" ;
	$ComposerServersCsvValidationCheck.Text = "" ;
	$ConnectionServersCsvValidationCheck.Text = "" ;
	$PoolsCsvValidationCheck.Text = "" ;
	$DesktopsCsvValidationCheck.Text = "" ;
	$RDSServersCsvValidationCheck.Text = "" ;
	$FarmsCsvValidationCheck.Text = "" ;
	$ApplicationsCsvValidationCheck.Text = "" ;
	$GatewaysCsvValidationCheck.Text = "" ;
	$CsvValidationButton.Forecolor = [System.Drawing.Color]::Black ;
	$CsvValidationButton.Text = "Check for CSVs" ;
	$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Black ;
	$VisioOpenOutputButton.Text = "Select Visio Output Folder" ;
	$Infrastructure_DrawCheckBox.CheckState = "Checked" ;
	$Infrastructure_Complete.Text = "" ;
	$Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$Pool_to_Desktop_Complete.Text = "" ;
	$LC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$LC_Pool_to_Desktop_Complete.Text = "" ;
	$LC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$LC_Pool_to_Application_Complete.Text = "" ;
	$FC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$FC_Pool_to_Desktop_Complete.Text = "" ;
	$FC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$FC_Pool_to_Application_Complete.Text = "" ;
	$IC_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$IC_Pool_to_Desktop_Complete.Text = "" ;
	$IC_Pool_to_Application_DrawCheckBox.CheckState = "Checked" ;
	$IC_Pool_to_Application_Complete.Text = "" ;
	$Unmanaged_Pool_to_Desktop_DrawCheckBox.CheckState = "Checked" ;
	$Unmanaged_Pool_to_Desktop_Complete.Text = "" ;
	$Farm_to_Remote_Desktop_Services_DrawCheckBox.CheckState = "Checked" ;
	$Farm_to_Remote_Desktop_Services_Complete.Text = "" ;
	$DrawButton.Forecolor = [System.Drawing.Color]::Black ;
	$DrawButton.Text = "Draw Visio"
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Opening drawing." -ForegroundColor Magenta
	}
	if ( $logdraw -eq $true ) `
	{ `
		Stop-Transcript
	}
} )
#endregion ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Event Loop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Main
{ `
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run( $vDiagram)
}
#endregion ~~< Event Loop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Event Handlers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< HVServer Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Connect_HVServer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Connect_HVServer
{ `
	$global:ConnServ = $ConnServTextBox.Text
	$User = $UserNameTextBox.Text
	$Host.UI.RawUI.WindowTitle = "vDiagram Horizon $MyVer connected to $ConnServ"
	
	$global:HorizonViewServer = Connect-HVServer -server $ConnServ -user $User -password $PasswordTextBox.Text
	$global:HorizonViewAPI = $HorizonViewServer.ExtensionData
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
}
#endregion ~~< Connect_HVServer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Disconnect_HVServer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Disconnect_HVServer
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Disconnect-HVServer -Confirm:$false
	$Host.UI.RawUI.WindowTitle = "vDiagram Horizon $MyVer connected from $ConnServ"
	if ( $debug -eq $true ) `
	{ `
		Write-Host "[$DateTime] Disconnected from $ConnServ successfully." -ForegroundColor Magenta
	}
	Write-Host "[$DateTime] Click Open CSV Output Folder to view CSVs or proceed to drawing Visio." -ForegroundColor Yellow
}
#endregion ~~< Disconnect_HVServer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< HVServer Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Folder Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_CaptureCsvFolder
{ `
	$CaptureCsvBrowseLoop = $True
	while ( $CaptureCsvBrowseLoop ) `
	{ `
		if ( $CaptureCsvBrowse.ShowDialog() -eq "OK" ) `
		{ `
			$CaptureCsvBrowseLoop = $False
		}
		else `
		{ `
			$CaptureCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show( "You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel,[System.Windows.Forms.MessageBoxIcon]::Question )
			if ( $CaptureCsvBrowseRes -eq "Cancel" ) `
			{ `
				return
			}
		}
	}
	$global:CaptureCsvFolder = $CaptureCsvBrowse.SelectedPath
}
#endregion ~~< Find_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Check_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Check_CaptureCsvFolder
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	$CheckContentPath = $CaptureCsvFolder + "\" + $ConnServTextBox.Text
	$CheckContentDir = $CheckContentPath + "*.csv"
	$CheckContent = Test-Path $CheckContentDir
	if ( $CheckContent -eq "True" ) `
	{
		$CheckContents_CaptureCsvFolder =  [System.Windows.Forms.MessageBox]::Show( "Files where found in the folder. Would you like to delete these files? Click 'Yes' to delete and 'No' move files to a new folder.","Warning!",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question )
		switch  ( $CheckContents_CaptureCsvFolder ) `
		{ `
			'Yes' 
			{ `
				Remove-Item $CheckContentDir
				if ( $debug -eq $true ) `
				{ `
					Write-Host "[$DateTime] Files were present in folder. Deleting files from folder."
				}
			}
			'No'
			{ `
				$CheckContentCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
				$CheckContentCsvBrowse.Description = "Select a directory to copy files to"
				$CheckContentCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
				$CheckContentCsvBrowse.ShowDialog()
				$global:NewContentCsvFolder = $CheckContentCsvBrowse.SelectedPath
				copy-item -Path $CheckContentDir -Destination $NewContentCsvFolder
				Remove-Item $CheckContentDir
				if ( $debug -eq $true ) `
				{ `
					Write-Host "[$DateTime] Files were present in folder. Moving old files to $NewContentCsvFolder"
				}
			}
		}
	}
}
#endregion ~~< Check_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_DrawCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_DrawCsvFolder
{ `
	$DrawCsvBrowseLoop = $True
	while ( $DrawCsvBrowseLoop ) `
	{ `
		if ( $DrawCsvBrowse.ShowDialog() -eq "OK" ) `
		{ `
			$DrawCsvBrowseLoop = $False
		}
		else `
		{
			$DrawCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show( "You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel )
			if ( $DrawCsvBrowseRes -eq "Cancel" ) `
			{ `
				return
			}
		}
	}
	$global:DrawCsvFolder = $DrawCsvBrowse.SelectedPath
}
#endregion ~~< Find_DrawCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_DrawVisioFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_DrawVisioFolder
{ `
	$VisioBrowseLoop = $True
	while( $VisioBrowseLoop ) `
	{ `
		if ( $VisioBrowse.ShowDialog() -eq "OK" ) `
		{ `
			$VisioBrowseLoop = $False
		}
		else `
		{ `
			$VisioBrowseRes = [System.Windows.Forms.MessageBox]::Show( "You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel )
			if( $VisioBrowseRes -eq "Cancel" ) `
			{ `
				return
			}
		}
	}
	$global:VisioFolder = $VisioBrowse.SelectedPath
}
#endregion ~~< Find_DrawVisioFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Folder Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Export Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VirtualCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VirtualCenter_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export vCenter Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting vCenter Info." -ForegroundColor Green
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	$VirtualCenter_CSV = "$CaptureCsvFolder\$ConnServ-VirtualCenterExport.csv"
	$VirtualCenter = $HorizonViewAPI.VirtualCenter.VirtualCenter_List()
	$VirtualCenterHealth = $HorizonViewAPI.VirtualCenterHealth.VirtualCenterHealth_List()
	$i = 0
	$VirtualCenterNumber = 0
	
	ForEach ( $VC in $VirtualCenterHealth ) `
	{  `
		if ( $debug -eq $true ) `
		{ `
			$VirtualCenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on vCenter object $VirtualCenterNumber of $( ( $VirtualCenterHealth ).Count ) -" ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.ServerName )
		}
	$i++
	$VirtualCenterCsvValidationComplete.Forecolor = "Blue"
	$VirtualCenterCsvValidationComplete.Text = "$i of $( $VirtualCenterHealth ).Count )"
		
	$VC | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "URL"; Expression = { [string]::Join( ", ", ( $_.Data.Name ) ) } }, `
			@{ Name = "Version"; Expression = { [string]::Join( ", ", ( $_.Data.Version ) ) } }, `
			@{ Name = "Build"; Expression = { [string]::Join( ", ", ( $_.Data.Build ) ) } }, `
			@{ Name = "ApiVersion"; Expression = { [string]::Join( ", ", ( $_.Data.ApiVersion ) ) } }, `
			@{ Name = "InstanceUuid"; Expression = { [string]::Join( ", ", ( $_.Data.InstanceUuid ) ) } }, `
			@{ Name = "ConnectionServerData_Id"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Id.Id ) ) } }, `
			@{ Name = "ConnectionServerData_Name"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Name ) ) } }, `
			@{ Name = "ConnectionServerData_Status"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Status ) ) } }, `
			@{ Name = "ConnectionServerData_ThumbprintAccepted"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.ThumbprintAccepted ) ) } }, `
			@{ Name = "ConnectionServerData_CertificateHealth_Valid"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.Valid ) ) } }, `
			@{ Name = "ConnectionServerData_CertificateHealth_StartTime"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.StartTime ) ) } }, `
			@{ Name = "ConnectionServerData_CertificateHealth_ExpirationTime"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.ExpirationTime ) ) } }, `
			@{ Name = "ConnectionServerData_CertificateHealth_InvalidReason"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.InvalidReason ) ) } }, `
			@{ Name = "ConnectionServerData_CertificateHealth_ConnectionServerCertificate"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.ConnectionServerCertificate ) ) } }, `
			@{ Name = "HostData_Name"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).Name ) ) } }, `
			@{ Name = "HostData_Version"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).Version ) ) } }, `
			@{ Name = "HostData_ApiVersion"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).ApiVersion ) ) } }, `
			@{ Name = "HostData_Status"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).Status ) ) } }, `
			@{ Name = "HostData_ClusterName"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).ClusterName ) ) } }, `
			@{ Name = "HostData_VGPUTypes"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).VGPUTypes ) ) } }, `
			@{ Name = "HostData_NumCpuCores"; Expression = { [string]::Join( ", ", ( $( $_.HostData | Sort-Object Name ).NumCpuCores ) ) } }, `
			@{ Name = "HostData_CpuMhz"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).CpuMhz ) ) } }, `
			@{ Name = "HostData_OverallCpuUsage"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).OverallCpuUsage ) ) } }, `
			@{ Name = "HostData_MemorySizeBytes"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).MemorySizeBytes ) ) } }, `
			@{ Name = "HostData_OverallMemoryUsageMB"; Expression = { [string]::Join( ", ", ( ( $_.HostData | Sort-Object Name ).OverallMemoryUsageMB ) ) } }, `
			@{ Name = "DatastoreData_Id_Id"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).Id.Id ) ) } }, `
			@{ Name = "DatastoreData_Name"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).Name ) ) } }, `
			@{ Name = "DatastoreData_Accessible"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).Accessible ) ) } }, `
			@{ Name = "DatastoreData_Path"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).Path ) ) } }, `
			@{ Name = "DatastoreData_DatastoreType"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).DatastoreType ) ) } }, `
			@{ Name = "DatastoreData_CapacityMB"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).CapacityMB ) ) } }, `
			@{ Name = "DatastoreData_FreeSpaceMB"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).FreeSpaceMB ) ) } }, `
			@{ Name = "DatastoreData_Url"; Expression = { [string]::Join( ", ", ( ( $_.DatastoreData | Sort-Object Name ).Url ) ) } },
			@{ Name = "ServerName"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.ServerName ) ) } }, `
			@{ Name = "Port"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.Port ) ) } }, `
			@{ Name = "UseSSL"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.UseSSL ) ) } }, `
			@{ Name = "UserName"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.UserName ) ) } }, `
			@{ Name = "ServerType"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ServerSpec.ServerType ) ) } }, `
			@{ Name = "Description"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Description ) ) } }, `
			@{ Name = "DisplayName"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).DisplayName ) ) } }, `
			@{ Name = "CertificateOverride"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).CertificateOverride ) ) } }, `
			@{ Name = "Limits_VcProvisioningLimit"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Limits.VcProvisioningLimit ) ) } }, `
			@{ Name = "Limits_VcPowerOperationsLimit"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Limits.VcPowerOperationsLimit ) ) } }, `
			@{ Name = "Limits_ViewComposerProvisioningLimit"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Limits.ViewComposerProvisioningLimit ) ) } }, `
			@{ Name = "Limits_ViewComposerMaintenanceLimit"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Limits.ViewComposerMaintenanceLimit ) ) } }, `
			@{ Name = "Limits_InstantCloneEngineProvisioningLimit"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Limits.InstantCloneEngineProvisioningLimit ) ) } }, `
			@{ Name = "StorageAcceleratorData_Enabled"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).StorageAcceleratorData.Enabled ) ) } }, `
			@{ Name = "StorageAcceleratorData_DefaultCacheSizeMB"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).StorageAcceleratorData.DefaultCacheSizeMB ) ) } }, `
			@{ Name = "StorageAcceleratorData_HostOverrides"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).StorageAcceleratorData.HostOverrides ) ) } }, `
			@{ Name = "ViewComposerData_ViewComposerType"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ViewComposerType ) ) } }, `
			@{ Name = "ViewComposerData_ServerSpec_ServerName"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ServerSpec.ServerName ) ) } }, `
			@{ Name = "ViewComposerData_ServerSpec_Port"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ServerSpec.Port ) ) } }, `
			@{ Name = "ViewComposerData_ServerSpec_UseSSL"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ServerSpec.UseSSL ) ) } }, `
			@{ Name = "ViewComposerData_ServerSpec_UserName"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ServerSpec.UserName ) ) } }, `
			@{ Name = "ViewComposerData_ServerSpec_ServerType"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).ViewComposerData.ServerSpec.ServerType ) ) } }, `
			@{ Name = "SeSparseReclamationEnabled"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).SeSparseReclamationEnabled ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).Enabled ) ) } }, `
			@{ Name = "VmcDeployment"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).VmcDeployment ) ) } }, `
			@{ Name = "IsDeletable"; Expression = { [string]::Join( ", ", ( ( $VirtualCenter | Where-Object { $_.Id.Id -eq $VC.Id.Id } ).IsDeletable ) ) } } | `
		Export-Csv $VirtualCenter_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< VirtualCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ComposerServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ComposerServers_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Composer Server Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Composer Server Info." -ForegroundColor Green
	$ComposerServers_CSV = "$CaptureCsvFolder\$ConnServ-ComposerServersExport.csv"
	$ComposerServers = $HorizonViewAPI.ViewComposerHealth.ViewComposerHealth_List()
	$i = 0
	$ComposerServersNumber = 0
	
	ForEach ( $ComposerServer in $ComposerServers) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$ComposerServersNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Composer Server object $ComposerServersNumber of $( ( $ComposerServers ).Count ) -" $ComposerServer.ServerName
		}
		
	$ComposerServer | `
		Select-Object `
			@{ Name = "ServerName"; Expression = { [string]::Join( ", ", ( $_.ServerName ) ) } }, `
			@{ Name = "Port"; Expression = { [string]::Join( ", ", ( $_.Port ) ) } }, `
			@{ Name = "VirtualCenters_Id"; Expression = { [string]::Join( ", ", ( $_.Data.VirtualCenters.Id ) ) } }, `
			@{ Name = "Version"; Expression = { [string]::Join( ", ", ( $_.Data.Version ) ) } }, `
			@{ Name = "Build"; Expression = { [string]::Join( ", ", ( $_.Data.Build ) ) } }, `
			@{ Name = "ApiVersion"; Expression = { [string]::Join( ", ", ( $_.Data.ApiVersion ) ) } }, `
			@{ Name = "MinVCVersion"; Expression = { [string]::Join( ", ", ( $_.Data.MinVCVersion ) ) } }, `
			@{ Name = "MinESXVersion"; Expression = { [string]::Join( ", ", ( $_.Data.MinESXVersion ) ) } }, `
			@{ Name = "ConnectionServer_Id"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Id.Id ) ) } },`
			@{ Name = "ConnectionServer_Name"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Name ) ) } }, `
			@{ Name = "ConnectionServer_Status"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.Status ) ) } }, `
			@{ Name = "ConnectionServer_ErrorMessage"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.ErrorMessage ) ) } }, `
			@{ Name = "ConnectionServer_ThumbprintAccepted"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.ThumbprintAccepted ) ) } }, `
			@{ Name = "ConnectionServer_CertificateHealth_Valid"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.Valid ) ) } }, `
			@{ Name = "ConnectionServer_CertificateHealth_StartTime"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.StartTime ) ) } }, `
			@{ Name = "ConnectionServer_CertificateHealth_ExpirationTime"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.ExpirationTime ) ) } }, `
			@{ Name = "ConnectionServer_CertificateHealth_InvalidReason"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.InvalidReason ) ) } }, `
			@{ Name = "ConnectionServer_CertificateHealth_ConnectionServerCertificate"; Expression = { [string]::Join( ", ", ( $_.ConnectionServerData.CertificateHealth.ConnectionServerCertificate ) ) } } | `
		Sort-Object ServerName | `
		Export-Csv $ComposerServers_CSV -Append -NoTypeInformation
	$i++
	$ComposerServersCsvValidationComplete.Forecolor = "Blue"
	$ComposerServersCsvValidationComplete.Text = "$i of $( $ComposerServers ).Count )"
	}
}
#endregion ~~< ComposerServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ConnectionServers_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Connection Server Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Connection Server Info." -ForegroundColor Green
	$ConnectionServers_CSV = "$CaptureCsvFolder\$ConnServ-ConnectionServersExport.csv"
	$ConnectionServerHealth = $HorizonViewAPI.ConnectionServerHealth.ConnectionServerHealth_List()
	$ConnectionServers = $HorizonViewAPI.ConnectionServer.ConnectionServer_List()
	
	$i = 0
	$ConnectionServersNumber = 0
	
	ForEach ( $ConnectionServer in $ConnectionServers ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$ConnectionServersNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Connection Server object $ConnectionServersNumber of $( ( $ConnectionServers ).Count ) -" $ConnectionServer.General.Name
		}
	$i++
	$ConnectionServersCsvValidationComplete.Forecolor = "Blue"
	$ConnectionServersCsvValidationComplete.Text = "$i of $( $ConnectionServers ).Count )"

	$ConnectionServer | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.General.Name ) ) } }, `
			@{ Name = "ServerAddress"; Expression = { [string]::Join( ", ", ( $_.General.ServerAddress ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( $_.General.Enabled ) ) } }, `
			@{ Name = "Tags"; Expression = { [string]::Join( ", ", ( $_.General.Tags ) ) } }, `
			@{ Name = "ExternalURL"; Expression = { [string]::Join( ", ", ( $_.General.ExternalURL ) ) } }, `
			@{ Name = "ExternalPCoIPURL"; Expression = { [string]::Join( ", ", ( $_.General.ExternalPCoIPURL ) ) } }, `
			@{ Name = "HasPCoIPGatewaySupport"; Expression = { [string]::Join( ", ", ( $_.General.HasPCoIPGatewaySupport ) ) } }, `
			@{ Name = "HasBlastGatewaySupport"; Expression = { [string]::Join( ", ", ( $_.General.HasBlastGatewaySupport ) ) } }, `
			@{ Name = "AuxillaryExternalPCoIPIPv4Address"; Expression = { [string]::Join( ", ", ( $_.General.AuxillaryExternalPCoIPIPv4Address ) ) } }, `
			@{ Name = "ExternalAppblastURL"; Expression = { [string]::Join( ", ", ( $_.General.ExternalAppblastURL ) ) } }, `
			@{ Name = "LocalConnectionServer"; Expression = { [string]::Join( ", ", ( $_.General.LocalConnectionServer ) ) } }, `
			@{ Name = "BypassTunnel"; Expression = { [string]::Join( ", ", ( $_.General.BypassTunnel ) ) } }, `
			@{ Name = "BypassPCoIPGateway"; Expression = { [string]::Join( ", ", ( $_.General.BypassPCoIPGateway ) ) } }, `
			@{ Name = "BypassAppBlastGateway"; Expression = { [string]::Join( ", ", ( $_.General.BypassAppBlastGateway ) ) } }, `
			@{ Name = "DirectHTMLABSG"; Expression = { [string]::Join( ", ", ( $_.General.DirectHTMLABSG ) ) } }, `
			@{ Name = "FullVersion"; Expression = { [string]::Join( ", ", ( $_.General.Version ) ) } }, `
			@{ Name = "IpMode"; Expression = { [string]::Join( ", ", ( $_.General.IpMode ) ) } }, `
			@{ Name = "FipsModeEnabled"; Expression = { [string]::Join( ", ", ( $_.General.FipsModeEnabled ) ) } }, `
			@{ Name = "Fqhn"; Expression = { [string]::Join( ", ", ( $_.General.Fqhn ) ) } }, `
			@{ Name = "SmartCardSupport"; Expression = { [string]::Join( ", ", ( $_.Authentication.SmartCardSupport ) ) } }, `
			@{ Name = "EnableSmartCardUserNameHint"; Expression = { [string]::Join( ", ", ( $_.Authentication.EnableSmartCardUserNameHint ) ) } }, `
			@{ Name = "LogoffWhenRemoveSmartCard"; Expression = { [string]::Join( ", ", ( $_.Authentication.LogoffWhenRemoveSmartCard ) ) } }, `
			@{ Name = "SmartCardSupportForAdmin"; Expression = { [string]::Join( ", ", ( $_.Authentication.SmartCardSupportForAdmin ) ) } }, `
			@{ Name = "RsaSecureIdConfig_SecureIdEnabled"; Expression = { [string]::Join( ", ", ( $_.Authentication.RsaSecureIdConfig.SecureIdEnabled ) ) } }, `
			@{ Name = "RsaSecureIdConfig_NameMapping"; Expression = { [string]::Join( ", ", ( $_.Authentication.RsaSecureIdConfig.NameMapping ) ) } }, `
			@{ Name = "RsaSecureIdConfig_ClearNodeSecret"; Expression = { [string]::Join( ", ", ( $_.Authentication.RsaSecureIdConfig.ClearNodeSecret ) ) } }, `
			@{ Name = "RsaSecureIdConfig_SecurityFileData"; Expression = { [string]::Join( ", ", ( $_.Authentication.RsaSecureIdConfig.SecurityFileData ) ) } }, `
			@{ Name = "RsaSecureIdConfig_SecurityFileUploaded"; Expression = { [string]::Join( ", ", ( $_.Authentication.RsaSecureIdConfig.SecurityFileUploaded ) ) } }, `
			@{ Name = "RadiusConfig_RadiusEnabled"; Expression = { [string]::Join( ", ", ( $_.Authentication.RadiusConfig.RadiusEnabled ) ) } }, `
			@{ Name = "RadiusConfig_RadiusAuthenticator"; Expression = { [string]::Join( ", ", ( $_.Authentication.RadiusConfig.RadiusAuthenticator ) ) } }, `
			@{ Name = "RadiusConfig_RadiusNameMapping"; Expression = { [string]::Join( ", ", ( $_.Authentication.RadiusConfig.RadiusNameMapping ) ) } }, `
			@{ Name = "RadiusConfig_RadiusSSO"; Expression = { [string]::Join( ", ", ( $_.Authentication.RadiusConfig.RadiusSSO ) ) } }, `
			@{ Name = "SamlConfig_SamlSupport"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.SamlSupport ) ) } }, `
			@{ Name = "SamlConfig_SamlAuthenticator_Id"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.SamlAuthenticator.Id ) ) } }, `
			@{ Name = "SamlConfig_SamlAuthenticators_Id"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.SamlAuthenticators.Id ) ) } }, `
			@{ Name = "SamlConfig_WorkspaceOneData_WorkspaceOneModeEnabled"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.WorkspaceOneData.WorkspaceOneModeEnabled ) ) } }, `
			@{ Name = "SamlConfig_WorkspaceOneData_WorkspaceOneHostName"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.WorkspaceOneData.WorkspaceOneHostName ) ) } }, `
			@{ Name = "SamlConfig_WorkspaceOneData_WorkspaceOneBlockOldClients"; Expression = { [string]::Join( ", ", ( $_.Authentication.SamlConfig.WorkspaceOneData.WorkspaceOneBlockOldClients ) ) } }, `
			@{ Name = "UnauthenticatedAccessConfig_Enabled"; Expression = { [string]::Join( ", ", ( $_.Authentication.UnauthenticatedAccessConfig.Enabled ) ) } }, `
			@{ Name = "UnauthenticatedAccessConfig_DefaultUser"; Expression = { [string]::Join( ", ", ( $_.Authentication.UnauthenticatedAccessConfig.DefaultUser ) ) } }, `
			@{ Name = "UnauthenticatedAccessConfig_UserIdleTimeout"; Expression = { [string]::Join( ", ", ( $_.Authentication.UnauthenticatedAccessConfig.UserIdleTimeout ) ) } }, `
			@{ Name = "UnauthenticatedAccessConfig_ClientPuzzleDifficulty"; Expression = { [string]::Join( ", ", ( $_.Authentication.UnauthenticatedAccessConfig.ClientPuzzleDifficulty ) ) } }, `
			@{ Name = "UnauthenticatedAccessConfig_BlockUnsupportedClients"; Expression = { [string]::Join( ", ", ( $_.Authentication.UnauthenticatedAccessConfig.BlockUnsupportedClients ) ) } }, `
			@{ Name = "LdapBackupFrequencyTime"; Expression = { [string]::Join( ", ", ( $_.Backup.LdapBackupFrequencyTime ) ) } }, `
			@{ Name = "LdapBackupMaxNumber"; Expression = { [string]::Join( ", ", ( $_.Backup.LdapBackupMaxNumber ) ) } }, `
			@{ Name = "LdapBackupFolder"; Expression = { [string]::Join( ", ", ( $_.Backup.LdapBackupFolder ) ) } }, `
			@{ Name = "LastLdapBackupTime"; Expression = { [string]::Join( ", ", ( $_.Backup.LastLdapBackupTime ) ) } }, `
			@{ Name = "LastLdapBackupStatus"; Expression = { [string]::Join( ", ", ( $_.Backup.LastLdapBackupStatus ) ) } }, `
			@{ Name = "IsBackupInProgress"; Expression = { [string]::Join( ", ", ( $_.Backup.IsBackupInProgress ) ) } }, `
			@{ Name = "LdapBackupTimeOffset"; Expression = { [string]::Join( ", ", ( $_.Backup.LdapBackupTimeOffset ) ) } }, `
			@{ Name = "SecurityServerPairing"; Expression = { [string]::Join( ", ", ( $_.SecurityServerPairing ) ) } }, `
			@{ Name = "MessageSecurity_MessageSecurityEnhancedModeSupported"; Expression = { [string]::Join( ", ", ( $_.MessageSecurity.MessageSecurityEnhancedModeSupported ) ) } }, `
			@{ Name = "MessageSecurity_RouterSslThumbprints"; Expression = { [string]::Join( ", ", ( $_.MessageSecurity.RouterSslThumbprints ) ) } }, `
			@{ Name = "MessageSecurity_MsgSecurityPublicKey"; Expression = { [string]::Join( ", ", ( $_.MessageSecurity.MsgSecurityPublicKey ) ) } }, `
			@{ Name = "Status"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).Status ) ) } }, `
			@{ Name = "Version"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).Version ) ) } }, `
			@{ Name = "Build"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).Build ) ) } }, `
			@{ Name = "ConnectionData_NumConnections"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumConnections ) ) } }, `
			@{ Name = "ConnectionData_NumConnectionsHigh"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumConnectionsHigh ) ) } }, `
			@{ Name = "ConnectionData_NumViewComposerConnections"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumViewComposerConnections ) ) } }, `
			@{ Name = "ConnectionData_NumViewComposerConnectionsHigh"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumViewComposerConnectionsHigh ) ) } }, `
			@{ Name = "ConnectionData_NumTunneledSessions"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumTunneledSessions ) ) } }, `
			@{ Name = "ConnectionData_NumPSGSessions"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).ConnectionData.NumPSGSession ) ) } }, `
			@{ Name = "DefaultCertificate"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).DefaultCertificate ) ) } }, `
			@{ Name = "CertificateHealth_Valid"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).CertificateHealth.Valid ) ) } }, `
			@{ Name = "CertificateHealth_StartTime"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).CertificateHealth.StartTime ) ) } }, `
			@{ Name = "CertificateHealth_ExpirationTime"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).CertificateHealth.ExpirationTime ) ) } }, `
			@{ Name = "CertificateHealth_InvalidReason"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).CertificateHealth.InvalidReason ) ) } }, `
			@{ Name = "CertificateHealth_ConnectionServerCertificate"; Expression = { [string]::Join( ", ", ( ( $ConnectionServerHealth | Where-Object { $_.Id.Id -eq $ConnectionServer.Id.Id } ).CertificateHealth.ConnectionServerCertificate ) ) } }	| `
		Sort-Object Name | `
		Export-Csv $ConnectionServers_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< ConnectionServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pools_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Pools_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Pool Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Pool Info." -ForegroundColor Green
	$Pools_CSV = "$CaptureCsvFolder\$ConnServ-PoolsExport.csv"
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"

	$Query.QueryEntityType = 'DesktopAssignmentView'
	$DesktopAssignmentViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$DesktopAssignmentView = $DesktopAssignmentViewQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$Query.QueryEntityType = 'DesktopSummaryView'
	$DesktopSummaryViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$DesktopSummaryView = $DesktopSummaryViewQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$i = 0
	$PoolNumber = 0
	
	ForEach	( $Pool in $DesktopSummaryView ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$PoolNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Desktop object $PoolNumber of $( ( $DesktopSummaryView ).Count ) -" $Pool.DesktopSummaryData.Name
		}
	$i++
	$PoolsCsvValidationComplete.Forecolor = "Blue"
	$PoolsCsvValidationComplete.Text = "$i of $( $DesktopSummaryView ).Count )"
		
	$Pool | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Name ) ) } }, `
			@{ Name = "DisplayName"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.DisplayName ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Enabled ) ) } }, `
			@{ Name = "Deleting"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Deleting ) ) } }, `
			@{ Name = "Type"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Type ) ) } }, `
			@{ Name = "Source"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Source ) ) } }, `
			@{ Name = "UserAssignment"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.UserAssignment ) ) } }, `
			@{ Name = "AccessGroup_Id"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.AccessGroup.Id ) ) } }, `
			@{ Name = "GlobalEntitlement"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.GlobalEntitlement ) ) } }, `
			@{ Name = "VirtualCenter_Id"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.VirtualCenter.Id ) ) } }, `
			@{ Name = "ProvisioningEnabled"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.ProvisioningEnabled ) ) } }, `
			@{ Name = "NumMachines"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.NumMachines ) ) } }, `
			@{ Name = "NumSessions"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.NumSessions ) ) } }, `
			@{ Name = "Farm_Id"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.Farm.Id ) ) } }, `
			@{ Name = "SupportedDomains"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.SupportedDomains ) ) } }, `
			@{ Name = "LastProvisioningError"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.LastProvisioningError ) ) } }, `
			@{ Name = "CategoryFolderName"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.CategoryFolderName ) ) } }, `
			@{ Name = "EnableAppRemoting"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.EnableAppRemoting ) ) } }, `
			@{ Name = "ApplicationCount"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.ApplicationCount ) ) } }, `
			@{ Name = "SupportedSessionType"; Expression = { [string]::Join( ", ", ( $_.DesktopSummaryData.SupportedSessionType ) ) } },	
			@{ Name = "OperatingSystem"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.OperatingSystem ) ) } }, `
			@{ Name = "OperatingSystemArchitecture"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.OperatingSystemArchitecture ) ) } }, `
			@{ Name = "EnableGRIDvGPUs"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.EnableGRIDvGPUs ) ) } }, `
			@{ Name = "Renderer3D"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.Renderer3D ) ) } }, `
			@{ Name = "AllowUsersToChooseProtocol"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.AllowUsersToChooseProtocol ) ) } }, `
			@{ Name = "AllowMultipleSessionsPerUser"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.AllowMultipleSessionsPerUser ) ) } }, `
			@{ Name = "AllowUsersToResetMachines"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.AllowUsersToResetMachines ) ) } }, `
			@{ Name = "DefaultDisplayProtocol"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.DefaultDisplayProtocol ) ) } }, `
			@{ Name = "EnableHTMLAccess"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.EnableHTMLAccess ) ) } }, `
			@{ Name = "EnableCollaboration"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.EnableCollaboration ) ) } }, `
			@{ Name = "MultipleSessionAutoClean"; Expression = { [string]::Join( ", ", ( ( $DesktopAssignmentView | Where-Object { $_.Id.Id -eq $Pool.Id.Id } ).DesktopAssignmentData.MultipleSessionAutoClean ) ) } } | `
		Sort-Object DSV_DesktopSummaryData_Name | `
		Export-Csv $Pools_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< Pools_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Desktops_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Desktops_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Desktop Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Desktop Info." -ForegroundColor Green
	$Desktops_CSV = "$CaptureCsvFolder\$ConnServ-DesktopsExport.csv"
	
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"
	$Query.limit = 1000
	$Query.maxpagesize = 1000

	$Query.QueryEntityType = 'MachineDetailsView'
	$MachineDetailsViewOffset = 0
	$MachineDetailsViewResults = @()
	do `
	{ `
		$Query.startingoffset = $MachineDetailsViewOffset
		$MachineDetailsViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
		if ( ( $MachineDetailsViewQuery.Results ).Count -eq 1000 ) `
		{ `
			$maxresults = 1
		} `
		else `
		{ `
			$maxresults = 0
		} `
		
		$MachineDetailsViewOffset += 1000
		$MachineDetailsViewResults += $MachineDetailsViewQuery.Results
	}
	until `
	( `
		$maxresults -eq 0
	)
	$MachineDetailsView = $MachineDetailsViewResults
	
	$Query.QueryEntityType = 'MachineStateView'
	$MachineStateViewOffset = 0
	$MachineStateViewResults = @()
	do `
	{ `
		$Query.startingoffset = $MachineStateViewOffset
		$MachineStateViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
		if ( ( $MachineStateViewQuery.Results).Count -eq 1000 ) `
		{ `
			$maxresults = 1
		}
		else `
		{ `
			$maxresults = 0
		} `
		
		$MachineStateViewOffset += 1000
		$MachineStateViewResults += $MachineStateViewQuery.Results
	}
	until `
	( `
		$maxresults -eq 0
	)
	$MachineStateView = $MachineStateViewResults
	
	$Query.QueryEntityType = 'MachineNamesView'
	$MachineNamesViewOffset = 0
	$MachineNamesViewResults = @()
	do `
	{ `
		$Query.startingoffset = $MachineNamesViewOffset
		$MachineNamesViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
		if ( ( $MachineNamesViewQuery.Results).Count -eq 1000 ) `
		{ `
			$maxresults = 1
		}
		else `
		{ `
			$maxresults = 0
		}
		
		$MachineNamesViewOffset += 1000
		$MachineNamesViewResults += $MachineNamesViewQuery.Results
	}
	until `
	( `
		$maxresults -eq 0
	)
	$MachineNamesView = $MachineNamesViewResults
	
	$i = 0
	$DesktopsNumber = 0
	
	ForEach ( $Desktop in $MachineNamesView ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$DesktopsNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Desktop object $DesktopsNumber of $( ( $MachineNamesView ).Count ) -" ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).Data.Name
		}
	$i++
	$DesktopsCsvValidationComplete.Forecolor = "Blue"
	$DesktopsCsvValidationComplete.Text = "$i of $( $MachineNamesView ).Count )"
		
	$Desktop | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "GroupId"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).Group.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).Data.Name ) ) } }, `
			@{ Name = "AssignedUser_Id"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).Data.AssignedUser.Id ) ) } }, `
			@{ Name = "AssignedUserName"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).Data.AssignedUserName ) ) } }, `
			@{ Name = "Type"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).DesktopData.Type ) ) } }, `
			@{ Name = "Source"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).DesktopData.Source ) ) } }, `
			@{ Name = "UserAssignment"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).DesktopData.UserAssignment ) ) } }, `
			@{ Name = "SessionProtocol"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).SessionData.SessionProtocol ) ) } }, `
			@{ Name = "VirtualCenter_Id"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.VirtualCenter.Id ) ) } }, `
			@{ Name = "VirtualDisks_Path"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData_VirtualDisks_Path ) ) } }, `
			@{ Name = "VirtualDisks_DatastorePath"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.VirtualDisks.DatastorePath ) ) } }, `
			@{ Name = "VirtualDisks_CapacityMB"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.VirtualDisks.CapacityMB ) ) } }, `
			@{ Name = "PersistentDisks"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.PersistentDisks ) ) } }, `
			@{ Name = "LastMaintenanceTime"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.LastMaintenanceTime ) ) } }, `
			@{ Name = "Operation"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.Operation ) ) } }, `
			@{ Name = "OperationState"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.OperationState ) ) } }, `
			@{ Name = "AutoRefreshLogOffSetting"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.AutoRefreshLogOffSetting ) ) } }, `
			@{ Name = "InHoldCustomization"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.InHoldCustomization ) ) } }, `
			@{ Name = "MissingInVCenter"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.MissingInVCenter ) ) } }, `
			@{ Name = "CreateTime"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.CreateTime ) ) } }, `
			@{ Name = "CloneErrorMessage"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.CloneErrorMessage ) ) } }, `
			@{ Name = "CloneErrorTime"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.CloneErrorTime ) ) } }, `
			@{ Name = "BaseImagePath"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.BaseImagePath ) ) } }, `
			@{ Name = "BaseImageSnapshotPath"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.BaseImageSnapshotPath ) ) } }, `
			@{ Name = "PendingBaseImagePath"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.PendingBaseImagePath ) ) } }, `
			@{ Name = "PendingBaseImageSnapshotPath"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).ManagedMachineDetailsData.PendingBaseImageSnapshotPath ) ) } }, `
			@{ Name = "PairingState"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).MachineAgentPairingData.PairingState ) ) } }, `
			@{ Name = "ConfiguredByBroker"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).MachineAgentPairingData.ConfiguredByBroker ) ) } }, `
			@{ Name = "AttemptedTheftByBroker"; Expression = { [string]::Join( ", ", ( ( $MachineDetailsView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).MachineAgentPairingData.AttemptedTheftByBroker ) ) } }, `
			@{ Name = "MachinePowerState"; Expression = { [string]::Join( ", ", ( ( $MachineStateView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).MachinePowerState ) ) } }, `
			@{ Name = "IpV4"; Expression = { [string]::Join( ", ", ( ( $MachineStateView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).IpV4 ) ) } }, `
			@{ Name = "IpV6"; Expression = { [string]::Join( ", ", ( ( $MachineStateView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).IpV6 ) ) } }, `
			@{ Name = "AgentId"; Expression = { [string]::Join( ", ", ( ( $MachineStateView | Where-Object { $_.Id.Id -eq $Desktop.Id.Id } ).AgentId ) ) } }, `
			@{ Name = "DnsName"; Expression = { [string]::Join( ", ", ( $_.Base.DnsName ) ) } }, `
			@{ Name = "User_Id"; Expression = { [string]::Join( ", ", ( $_.Base.User.Id ) ) } }, `
			@{ Name = "AccessGroup_Id"; Expression = { [string]::Join( ", ", ( $_.Base.AccessGroup.Id ) ) } }, `
			@{ Name = "Desktop_Id"; Expression = { [string]::Join( ", ", ( $_.Base.Desktop.Id ) ) } }, `
			@{ Name = "DesktopName"; Expression = { [string]::Join( ", ", ( $_.Base.DesktopName ) ) } }, `
			@{ Name = "Session_Id"; Expression = { [string]::Join( ", ", ( $_.Base.Session.Id ) ) } }, `
			@{ Name = "BasicState"; Expression = { [string]::Join( ", ", ( $_.Base.BasicState ) ) } }, `
			@{ Name = "Base_Type"; Expression = { [string]::Join( ", ", ( $_.Base.Type ) ) } }, `
			@{ Name = "OperatingSystem"; Expression = { [string]::Join( ", ", ( $_.Base.OperatingSystem ) ) } }, `
			@{ Name = "OperatingSystemArchitecture"; Expression = { [string]::Join( ", ", ( $_.Base.OperatingSystemArchitecture ) ) } }, `
			@{ Name = "AgentVersion"; Expression = { [string]::Join( ", ", ( $_.Base.AgentVersion ) ) } }, `
			@{ Name = "AgentBuildNumber"; Expression = { [string]::Join( ", ", ( $_.Base.AgentBuildNumber ) ) } }, `
			@{ Name = "RemoteExperienceAgentVersion"; Expression = { [string]::Join( ", ", ( $_.Base.RemoteExperienceAgentVersion ) ) } }, `
			@{ Name = "RemoteExperienceAgentBuildNumber"; Expression = { [string]::Join( ", ", ( $_.Base.RemoteExperienceAgentBuildNumber ) ) } }, `
			@{ Name = "UserName"; Expression = { [string]::Join( ", ", ( $_.NamesData.UserName ) ) } }, `
			@{ Name = "MessageSecurityMode"; Expression = { [string]::Join( ", ", ( $_.MessageSecurityData.MessageSecurityMode ) ) } }, `
			@{ Name = "MessageSecurityEnhancedModeSupported"; Expression = { [string]::Join( ", ", ( $_.MessageSecurityData.MessageSecurityEnhancedModeSupported ) ) } }, `
			@{ Name = "HostName"; Expression = { [string]::Join( ", ", ( $_.ManagedMachineNamesData.HostName ) ) } }, `
			@{ Name = "DatastorePaths"; Expression = { [string]::Join( ", ", ( $_.ManagedMachineNamesData.DatastorePaths ) ) } } | `
		Sort-Object DnsName | `
		Export-Csv $Desktops_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< Desktops_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDSServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function RDSServers_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export RDS Server Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting RDS Server Info." -ForegroundColor Green
	$RDSServers_CSV = "$CaptureCsvFolder\$ConnServ-RDSServersExport.csv"
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"

	$Query.QueryEntityType = 'RDSServerStateView'
	$RDSServerStateViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$RDSServerStateView = $RDSServerStateViewQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$Query.QueryEntityType = 'RDSServerSummaryView'
	$RDSServerSummaryViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$RDSServerSummaryView = $RDSServerSummaryViewQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$Query.QueryEntityType = 'RDSServerInfo'
	$RDSServerInfoQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$RDSServerInfo = $RDSServerInfoQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$i = 0
	$RDSServersNumber = 0
	
	ForEach ( $RDSServer in $RDSServerInfo ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$RDSServersNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on RDS Server object $RDSServersNumber of $( ( $RDSServerInfo ).Count ) -" $RDSServer.Base.Name
		}
	$i++
	$RDSServersCsvValidationComplete.Forecolor = "Blue"
	$RDSServersCsvValidationComplete.Text = "$i of $( $RDSServerInfo ).Count )"
		
	$RDSServer | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.Base.Name ) ) } }, `
			@{ Name = "Description"; Expression = { [string]::Join( ", ", ( $_.Base.Description ) ) } }, `
			@{ Name = "Farm_Id"; Expression = { [string]::Join( ", ", ( $_.Base.Farm.Id ) ) } }, `
			@{ Name = "Desktop_Id"; Expression = { [string]::Join( ", ", ( $_.Base.Desktop.Id ) ) } }, `
			@{ Name = "AccessGroup_Id"; Expression = { [string]::Join( ", ", ( $_.Base.AccessGroup.Id ) ) } }, `
			@{ Name = "MessageSecurityMode"; Expression = { [string]::Join( ", ", ( $_.MessageSecurityData.MessageSecurityMode ) ) } }, `
			@{ Name = "MessageSecurityEnhancedModeSupported"; Expression = { [string]::Join( ", ", ( $_.MessageSecurityData.MessageSecurityEnhancedModeSupported ) ) } }, `
			@{ Name = "DnsName"; Expression = { [string]::Join( ", ", ( $_.AgentData.DnsName ) ) } }, `
			@{ Name = "OperatingSystem"; Expression = { [string]::Join( ", ", ( $_.AgentData.OperatingSystem ) ) } }, `
			@{ Name = "AgentVersion"; Expression = { [string]::Join( ", ", ( $_.AgentData.AgentVersion ) ) } }, `
			@{ Name = "AgentBuildNumber"; Expression = { [string]::Join( ", ", ( $_.AgentData.AgentBuildNumber ) ) } }, `
			@{ Name = "RemoteExperienceAgentVersion"; Expression = { [string]::Join( ", ", ( $_.AgentData.RemoteExperienceAgentVersion ) ) } }, `
			@{ Name = "RemoteExperienceAgentBuildNumber"; Expression = { [string]::Join( ", ", ( $_.AgentData.RemoteExperienceAgentBuildNumber ) ) } }, `
			@{ Name = "SessionSettings_MaxSessionsType"; Expression = { [string]::Join( ", ", ( $_.Settings.SessionSettings.MaxSessionsType ) ) } }, `
			@{ Name = "SessionSettings_MaxSessionsSetByAdmin"; Expression = { [string]::Join( ", ", ( $_.Settings.SessionSettings.MaxSessionsSetByAdmin ) ) } }, `
			@{ Name = "Agent_MaxSessionsType"; Expression = { [string]::Join( ", ", ( $_.Settings.AgentMaxSessionsData.MaxSessionsType ) ) } }, `
			@{ Name = "Agent_MaxSessionsSetByAdmin"; Expression = { [string]::Join( ", ", ( $_.Settings.AgentMaxSessionsData.MaxSessionsSetByAdmin ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( $_.Settings.Enabled ) ) } }, `
			@{ Name = "Status"; Expression = { [string]::Join( ", ", ( $_.RuntimeData.Status ) ) } }, `
			@{ Name = "SessionCount"; Expression = { [string]::Join( ", ", ( $_.RuntimeData.SessionCount ) ) } }, `
			@{ Name = "LoadPreference"; Expression = { [string]::Join( ", ", ( $_.RuntimeData.LoadPreference ) ) } }, `
			@{ Name = "LoadIndex"; Expression = { [string]::Join( ", ", ( $_.RuntimeData.LoadIndex ) ) } }, `
			@{ Name = "Operation"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.Operation ) ) } }, `
			@{ Name = "OperationState"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.OperationState ) ) } }, `
			@{ Name = "LogOffSetting"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.LogOffSetting ) ) } }, `
			@{ Name = "BaseImagePath"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.BaseImagePath ) ) } }, `
			@{ Name = "BaseImageSnapshotPath"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.BaseImageSnapshotPath ) ) } }, `
			@{ Name = "PendingBaseImagePath"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.PendingBaseImagePath ) ) } }, `
			@{ Name = "PendingBaseImageSnapshotPath"; Expression = { [string]::Join( ", ", ( $_.RdsServerMaintenanceData.PendingBaseImageSnapshotPath ) ) } },
			@{ Name = "FarmName"; Expression = { [string]::Join( ", ", ( ( $RDSServerSummaryView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).SummaryData.FarmName ) ) } }, `
			@{ Name = "DesktopName"; Expression = { [string]::Join( ", ", ( ( $RDSServerSummaryView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).SummaryData.DesktopName ) ) } }, `
			@{ Name = "FarmType"; Expression = { [string]::Join( ", ", ( ( $RDSServerSummaryView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).SummaryData.FarmType ) ) } }, `
			@{ Name = "MachinePowerState"; Expression = { [string]::Join( ", ", ( ( $RDSServerStateView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).MachinePowerState ) ) } }, `
			@{ Name = "IpV4"; Expression = { [string]::Join( ", ", ( ( $RDSServerStateView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).IpV4 ) ) } }, `
			@{ Name = "IpV6"; Expression = { [string]::Join( ", ", ( ( $RDSServerStateView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).IpV6 ) ) } }, `
			@{ Name = "AgentId"; Expression = { [string]::Join( ", ", ( ( $RDSServerStateView | Where-Object { $_.Id.Id -eq $RDSServer.Id.Id } ).AgentId ) ) } } | `
		Export-Csv $RDSServers_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< RDSServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farms_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Farms_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Farm Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Farm Info." -ForegroundColor Green
	$Farms_CSV = "$CaptureCsvFolder\$ConnServ-FarmsExport.csv"
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"

	$Query.QueryEntityType = 'FarmSummaryView'
	$FarmSummaryViewQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$FarmSummaryView = $FarmSummaryViewQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$Query.QueryEntityType = 'FarmHealthInfo'
	$FarmHealthInfoQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$FarmHealthInfo = $FarmHealthInfoQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
	
	$i = 0
	$FarmsNumber = 0
	
	ForEach ( $Farm in $FarmHealthInfo ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$FarmsNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Farms object $FarmsNumber of $( ( $FarmHealthInfo ).Count ) -" $Farm.Name
		}
	$i++
	$FarmsCsvValidationComplete.Forecolor = "Blue"
	$FarmsCsvValidationComplete.Text = "$i of $( $FarmHealthInfo ).Count )"
		
	$Farm | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.Name ) ) } }, `
			@{ Name = "Type"; Expression = { [string]::Join( ", ", ( $_.Type ) ) } }, `
			@{ Name = "Health"; Expression = { [string]::Join( ", ", ( $_.Health ) ) } }, `
			@{ Name = "Source"; Expression = { [string]::Join( ", ", ( $_.Source ) ) } }, `
			@{ Name = "AccessGroup_Id"; Expression = { [string]::Join( ", ", ( $_.AccessGroup.Id ) ) } }, `
			@{ Name = "RdsServer_Id"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.Id.Id ) ) } }, `
			@{ Name = "RdsServer_Name"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.Name ) ) } }, `
			@{ Name = "RdsServer_OperatingSystem"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.OperatingSystem ) ) } }, `
			@{ Name = "RdsServer_AgentVersion"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.AgentVersion ) ) } }, `
			@{ Name = "RdsServer_AgentBuildNumber"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.AgentBuildNumber ) ) } }, `
			@{ Name = "RdsServer_Status"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.Status ) ) } }, `
			@{ Name = "RdsServer_Health"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.Health ) ) } }, `
			@{ Name = "RdsServer_Available"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.Available ) ) } }, `
			@{ Name = "RdsServer_MissingApplications"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.MissingApplications ) ) } }, `
			@{ Name = "RdsServer_LoadPreference"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.LoadPreference ) ) } }, `
			@{ Name = "RdsServer_LoadIndex"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.LoadIndex ) ) } }, `
			@{ Name = "RdsServer_SessionSettings_MaxSessionsType"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.SessionSettings.MaxSessionsType ) ) } }, `
			@{ Name = "RdsServer_SessionSettings_MaxSessionsSetByAdmin"; Expression = { [string]::Join( ", ", ( $_.RdsServerHealth.SessionSettings.MaxSessionsSetByAdmin ) ) } }, `
			@{ Name = "NumApplications"; Expression = { [string]::Join( ", ", ( $_.NumApplications ) ) } },
			@{ Name = "DisplayName"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.DisplayName ) ) } }, `
			@{ Name = "Description"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.Description ) ) } }, `
			@{ Name = "AccessGroupName"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.AccessGroupName ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.Enabled ) ) } }, `
			@{ Name = "ProvisioningEnabled"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.ProvisioningEnabled ) ) } }, `
			@{ Name = "Deleting"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.Deleting ) ) } }, `
			@{ Name = "Desktop_Id"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.Desktop.Id ) ) } }, `
			@{ Name = "DesktopName"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.DesktopName ) ) } }, `
			@{ Name = "RdsServerCount"; Expression = { [string]::Join( ", ", ( ( $FarmSummaryView | Where-Object { $_.Id.Id -eq $Farm.Id.Id } ).Data.RdsServerCount ) ) } } | `
		Export-Csv $Farms_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< Farms_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Applications_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Applications_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Application Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Application Info." -ForegroundColor Green
	$Applications_CSV = "$CaptureCsvFolder\$ConnServ-ApplicationsExport.csv"
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"

	$Query.QueryEntityType = 'ApplicationInfo'
	$ApplicationInfoQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$ApplicationInfo = $ApplicationInfoQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )

	$i = 0
	$ApplicationsNumber = 0
		
	ForEach ( $Application in $ApplicationInfo ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$ApplicationsNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Application object $ApplicationsNumber of $( ( $ApplicationInfo ).Count ) -" $Application.Data.Name
		}
	$i++
	$ApplicationsCsvValidationComplete.Forecolor = "Blue"
	$ApplicationsCsvValidationComplete.Text = "$i of $( $ApplicationInfo ).Count )"
		
	$Application | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.Data.Name ) ) } }, `
			@{ Name = "DisplayName"; Expression = { [string]::Join( ", ", ( $_.Data.DisplayName ) ) } }, `
			@{ Name = "Description"; Expression = { [string]::Join( ", ", ( $_.Data.Description ) ) } }, `
			@{ Name = "Enabled"; Expression = { [string]::Join( ", ", ( $_.Data.Enabled ) ) } }, `
			@{ Name = "GlobalApplicationEntitlement"; Expression = { [string]::Join( ", ", ( $_.Data.GlobalApplicationEntitlement ) ) } }, `
			@{ Name = "EnableAntiAffinityRules"; Expression = { [string]::Join( ", ", ( $_.Data.EnableAntiAffinityRules ) ) } }, `
			@{ Name = "AntiAffinityPatterns"; Expression = { [string]::Join( ", ", ( $_.Data.AntiAffinityPatterns ) ) } }, `
			@{ Name = "AntiAffinityCount"; Expression = { [string]::Join( ", ", ( $_.Data.AntiAffinityCount ) ) } }, `
			@{ Name = "EnablePreLaunch"; Expression = { [string]::Join( ", ", ( $_.Data.EnablePreLaunch ) ) } }, `
			@{ Name = "ConnectionServerRestrictions"; Expression = { [string]::Join( ", ", ( $_.Data.ConnectionServerRestrictions ) ) } }, `
			@{ Name = "CategoryFolderName"; Expression = { [string]::Join( ", ", ( $_.Data.CategoryFolderName ) ) } }, `
			@{ Name = "ClientRestrictions"; Expression = { [string]::Join( ", ", ( $_.Data.ClientRestrictions ) ) } }, `
			@{ Name = "ShortcutLocations"; Expression = { [string]::Join( ", ", ( $_.Data.ShortcutLocations ) ) } }, `
			@{ Name = "MultiSessionMode"; Expression = { [string]::Join( ", ", ( $_.Data.MultiSessionMode ) ) } }, `
			@{ Name = "MaxMultiSessions"; Expression = { [string]::Join( ", ", ( $_.Data.MaxMultiSessions ) ) } }, `
			@{ Name = "ExecutablePath"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.ExecutablePath ) ) } }, `
			@{ Name = "Version"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.Version ) ) } }, `
			@{ Name = "Publisher"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.Publisher ) ) } }, `
			@{ Name = "StartFolder"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.StartFolder ) ) } }, `
			@{ Name = "Args"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.Args ) ) } }, `
			@{ Name = "Farm_Id"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.Farm.Id ) ) } }, `
			@{ Name = "Desktop_Id"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.Desktop.Id ) ) } }, `
			@{ Name = "FileTypes_FileType"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.FileTypes.FileType ) ) } }, `
			@{ Name = "FileTypes_Description"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.FileTypes.Description ) ) } }, `
			@{ Name = "AutoUpdateFileTypes"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.AutoUpdateFileTypes ) ) } }, `
			@{ Name = "OtherFileTypes"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.OtherFileTypes ) ) } }, `
			@{ Name = "AutoUpdateOtherFileTypes"; Expression = { [string]::Join( ", ", ( $_.ExecutionData.AutoUpdateOtherFileTypes ) ) } }, `
			@{ Name = "Icons_Id"; Expression = { [string]::Join( ", ", ( $_.Icons.Id ) ) } }, `
			@{ Name = "CustomizedIcons_Id"; Expression = { [string]::Join( ", ", ( $_.CustomizedIcons.Id ) ) } }, `
			@{ Name = "AccessGroup_Id"; Expression = { [string]::Join( ", ", ( $_.AccessGroup.Id ) ) } } | `
		Sort-Object Name | `
		Export-Csv $Applications_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< Applications_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Gateways_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Gateways_Export
{ `
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Gateway Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Gateway Info." -ForegroundColor Green
	$Gateways_CSV = "$CaptureCsvFolder\$ConnServ-GatewaysExport.csv"
	$Gateways = $HorizonViewAPI.GatewayHealth.GatewayHealth_List()
	
	$i = 0
	$GatewaysNumber = 0
	
	ForEach ( $Gateway in $Gateways ) `
	{ `
		if ( $debug -eq $true ) `
		{ `
			$GatewaysNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Gateway object $GatewaysNumber of $( ( $Gateways ).Count ) -" $Gateway.Name
		}
	$i++
	$GatewaysCsvValidationComplete.Forecolor = "Blue"
	$GatewaysCsvValidationComplete.Text = "$i of $( $Gateways ).Count )"
		
	$Gateway | `
		Select-Object `
			@{ Name = "Id"; Expression = { [string]::Join( ", ", ( $_.Id.Id ) ) } }, `
			@{ Name = "Name"; Expression = { [string]::Join( ", ", ( $_.Name ) ) } }, `
			@{ Name = "Address"; Expression = { [string]::Join( ", ", ( $_.Address ) ) } }, `
			@{ Name = "GatewayZoneInternal"; Expression = { [string]::Join( ", ", ( $_.GatewayZoneInternal ) ) } }, `
			@{ Name = "Version"; Expression = { [string]::Join( ", ", ( $_.Version ) ) } }, `
			@{ Name = "Type"; Expression = { [string]::Join( ", ", ( $_.Type ) ) } }, `
			@{ Name = "ConnectionData_NumActiveConnections"; Expression = { [string]::Join( ", ", ( $_.ConnectionData.NumActiveConnections ) ) } }, `
			@{ Name = "ConnectionData_NumPcoipConnections"; Expression = { [string]::Join( ", ", ( $_.ConnectionData.NumPcoipConnections ) ) } }, `
			@{ Name = "ConnectionData_NumBlastConnections"; Expression = { [string]::Join( ", ", ( $_.ConnectionData.NumBlastConnections ) ) } }, `
			@{ Name = "GatewayStatusActive"; Expression = { [string]::Join( ", ", ( $_.GatewayStatusActive ) ) } }, `
			@{ Name = "GatewayStatusStale"; Expression = { [string]::Join( ", ", ( $_.GatewayStatusStale ) ) } }, `
			@{ Name = "GatewayContacted"; Expression = { [string]::Join( ", ", ( $_.GatewayContacted ) ) } } | `
		Sort-Object Name | `
		Export-Csv $Gateways_CSV -Append -NoTypeInformation
	}
}
#endregion ~~< Gateway_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Export Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Object Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Connect-VisioObject >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Connect-VisioObject( $firstObj, $secondObj )
{ `
	$shpConn = $pagObj.Drop( $pagObj.Application.ConnectorToolDataObject, 0, 0 )
	$ConnectBegin = $shpConn.CellsU( "BeginX" ).GlueTo( $firstObj.CellsU( "PinX" ) )
	$ConnectEnd = $shpConn.CellsU( "EndX" ).GlueTo( $secondObj.CellsU( "PinX" ) )
}
#endregion ~~< Connect-VisioObject >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectVirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVirtualCenter( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.ServerName
	return $shpObj
}
#endregion ~~< Add-VisioObjectVirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectComposer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectComposer( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.ServerName
	return $shpObj
}
#endregion ~~< Add-VisioObjectComposer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectConnection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectConnection( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.Name
	return $shpObj
}
#endregion ~~< Add-VisioObjectConnection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectPools >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectPools( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.DisplayName
	return $shpObj
}
#endregion ~~< Add-VisioObjectPools >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectDesktops >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDesktops( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.Name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDesktops >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectRDSH >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectRDSH( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.Name
	return $shpObj
}
#endregion ~~< Add-VisioObjectRDSH >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectFarm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectFarm( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.DisplayName
	return $shpObj
}
#endregion ~~< Add-VisioObjectFarm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectApplication >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectApplication( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.DisplayName
	return $shpObj
}
#endregion ~~< Add-VisioObjectApplication >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Add-VisioObjectGateway >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectGateway( $mastObj, $item )
{ `
	$shpObj = $pagObj.Drop( $mastObj, $x, $y )
	$shpObj.Text = $item.Name
	return $shpObj
}
#endregion ~~< Add-VisioObjectGateway >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Object Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Draw Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VirtualCenter
{ `
	$VirtualCenterObject.Cells( "Prop.Id" ).Formula = '"' + $VirtualCenter.Id + '"'
	$VirtualCenterObject.Cells( "Prop.Name" ).Formula = '"' + $VirtualCenter.Name + '"'
	$VirtualCenterObject.Cells( "Prop.Version" ).Formula = '"' + $VirtualCenter.Version + '"'
	$VirtualCenterObject.Cells( "Prop.Build" ).Formula = '"' + $VirtualCenter.Build + '"'
	$VirtualCenterObject.Cells( "Prop.ApiVersion" ).Formula = '"' + $VirtualCenter.ApiVersion + '"'
	$VirtualCenterObject.Cells( "Prop.InstanceUuid" ).Formula = '"' + $VirtualCenter.InstanceUuid + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_Id" ).Formula = '"' + $VirtualCenter.ConnectionServerData_Id + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_Name" ).Formula = '"' + $VirtualCenter.ConnectionServerData_Name + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_Status" ).Formula = '"' + $VirtualCenter.ConnectionServerData_Status + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_ThumbprintAccepted" ).Formula = '"' + $VirtualCenter.ConnectionServerData_ThumbprintAccepted + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_CertificateHealth_Valid" ).Formula = '"' + $VirtualCenter.ConnectionServerData_CertificateHealth_Valid + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_CertificateHealth_StartTime" ).Formula = '"' + $VirtualCenter.ConnectionServerData_CertificateHealth_StartTime + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_CertificateHealth_ExpirationTime" ).Formula = '"' + $VirtualCenter.ConnectionServerData_CertificateHealth_ExpirationTime + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_CertificateHealth_InvalidReason" ).Formula = '"' + $VirtualCenter.ConnectionServerData_CertificateHealth_InvalidReason + '"'
	$VirtualCenterObject.Cells( "Prop.ConnectionServerData_CertificateHealth_ConnectionServerCertificate" ).Formula = '"' + $VirtualCenter.ConnectionServerData_CertificateHealth_ConnectionServerCertificate + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_Name" ).Formula = '"' + $VirtualCenter.HostData_Name + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_Version" ).Formula = '"' + $VirtualCenter.HostData_Version + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_ApiVersion" ).Formula = '"' + $VirtualCenter.HostData_ApiVersion + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_Status" ).Formula = '"' + $VirtualCenter.HostData_Status + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_ClusterName" ).Formula = '"' + $VirtualCenter.HostData_ClusterName + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_VGPUTypes" ).Formula = '"' + $VirtualCenter.HostData_VGPUTypes + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_NumCpuCores" ).Formula = '"' + $VirtualCenter.HostData_NumCpuCores + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_CpuMhz" ).Formula = '"' + $VirtualCenter.HostData_CpuMhz + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_OverallCpuUsage" ).Formula = '"' + $VirtualCenter.HostData_OverallCpuUsage + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_MemorySizeBytes" ).Formula = '"' + $VirtualCenter.HostData_MemorySizeBytes + '"'
	$VirtualCenterObject.Cells( "Prop.HostData_OverallMemoryUsageMB" ).Formula = '"' + $VirtualCenter.HostData_OverallMemoryUsageMB + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_Id_Id" ).Formula = '"' + $VirtualCenter.DatastoreData_Id_Id + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_Name" ).Formula = '"' + $VirtualCenter.DatastoreData_Name + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_Accessible" ).Formula = '"' + $VirtualCenter.DatastoreData_Accessible + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_Path" ).Formula = '"' + $VirtualCenter.DatastoreData_Path + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_DatastoreType" ).Formula = '"' + $VirtualCenter.DatastoreData_DatastoreType + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_CapacityMB" ).Formula = '"' + $VirtualCenter.DatastoreData_CapacityMB + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_FreeSpaceMB" ).Formula = '"' + $VirtualCenter.DatastoreData_FreeSpaceMB + '"'
	$VirtualCenterObject.Cells( "Prop.DatastoreData_Url" ).Formula = '"' + $VirtualCenter.DatastoreData_Url + '"'
	$VirtualCenterObject.Cells( "Prop.ServerName" ).Formula = '"' + $VirtualCenter.ServerName + '"'
	$VirtualCenterObject.Cells( "Prop.Port" ).Formula = '"' + $VirtualCenter.Port + '"'
	$VirtualCenterObject.Cells( "Prop.UseSSL" ).Formula = '"' + $VirtualCenter.UseSSL + '"'
	$VirtualCenterObject.Cells( "Prop.UserName" ).Formula = '"' + $VirtualCenter.UserName + '"'
	$VirtualCenterObject.Cells( "Prop.ServerType" ).Formula = '"' + $VirtualCenter.ServerType + '"'
	$VirtualCenterObject.Cells( "Prop.Description" ).Formula = '"' + $VirtualCenter.Description + '"'
	$VirtualCenterObject.Cells( "Prop.DisplayName" ).Formula = '"' + $VirtualCenter.DisplayName + '"'
	$VirtualCenterObject.Cells( "Prop.CertificateOverride" ).Formula = '"' + $VirtualCenter.CertificateOverride + '"'
	$VirtualCenterObject.Cells( "Prop.Limits_VcProvisioningLimit" ).Formula = '"' + $VirtualCenter.Limits_VcProvisioningLimit + '"'
	$VirtualCenterObject.Cells( "Prop.Limits_VcPowerOperationsLimit" ).Formula = '"' + $VirtualCenter.Limits_VcPowerOperationsLimit + '"'
	$VirtualCenterObject.Cells( "Prop.Limits_ViewComposerProvisioningLimit" ).Formula = '"' + $VirtualCenter.Limits_ViewComposerProvisioningLimit + '"'
	$VirtualCenterObject.Cells( "Prop.Limits_ViewComposerMaintenanceLimit" ).Formula = '"' + $VirtualCenter.Limits_ViewComposerMaintenanceLimit + '"'
	$VirtualCenterObject.Cells( "Prop.Limits_InstantCloneEngineProvisioningLimit" ).Formula = '"' + $VirtualCenter.Limits_InstantCloneEngineProvisioningLimit + '"'
	$VirtualCenterObject.Cells( "Prop.StorageAcceleratorData_Enabled" ).Formula = '"' + $VirtualCenter.StorageAcceleratorData_Enabled + '"'
	$VirtualCenterObject.Cells( "Prop.StorageAcceleratorData_DefaultCacheSizeMB" ).Formula = '"' + $VirtualCenter.StorageAcceleratorData_DefaultCacheSizeMB + '"'
	$VirtualCenterObject.Cells( "Prop.StorageAcceleratorData_HostOverrides" ).Formula = '"' + $VirtualCenter.StorageAcceleratorData_HostOverrides + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ViewComposerType" ).Formula = '"' + $VirtualCenter.ViewComposerData_ViewComposerType + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ServerSpec_ServerName" ).Formula = '"' + $VirtualCenter.ViewComposerData_ServerSpec_ServerName + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ServerSpec_Port" ).Formula = '"' + $VirtualCenter.ViewComposerData_ServerSpec_Port + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ServerSpec_UseSSL" ).Formula = '"' + $VirtualCenter.ViewComposerData_ServerSpec_UseSSL + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ServerSpec_UserName" ).Formula = '"' + $VirtualCenter.ViewComposerData_ServerSpec_UserName + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_ServerSpec_ServerType" ).Formula = '"' + $VirtualCenter.ViewComposerData_ServerSpec_ServerType + '"'
	$VirtualCenterObject.Cells( "Prop.ViewComposerData_CertificateOverride" ).Formula = '"' + $VirtualCenter.ViewComposerData_CertificateOverride + '"'
	$VirtualCenterObject.Cells( "Prop.SeSparseReclamationEnabled" ).Formula = '"' + $VirtualCenter.SeSparseReclamationEnabled + '"'
	$VirtualCenterObject.Cells( "Prop.Enabled" ).Formula = '"' + $VirtualCenter.Enabled + '"'
	$VirtualCenterObject.Cells( "Prop.VmcDeployment" ).Formula = '"' + $VirtualCenter.VmcDeployment + '"'
	$VirtualCenterObject.Cells( "Prop.IsDeletable" ).Formula = '"' + $VirtualCenter.IsDeletable + '"'
}
#endregion ~~< Draw_VirtualCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Composer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Composer
{ `
	$ComposerObject.Cells( "Prop.ServerName" ).Formula = '"' + $ComposerServersImport.ServerName + '"'
	$ComposerObject.Cells( "Prop.Port" ).Formula = '"' + $ComposerServersImport.Port + '"'
	$ComposerObject.Cells( "Prop.VirtualCenters_Id" ).Formula = '"' + $ComposerServersImport.VirtualCenters_Id + '"'
	$ComposerObject.Cells( "Prop.Version" ).Formula = '"' + $ComposerServersImport.Version + '"'
	$ComposerObject.Cells( "Prop.Build" ).Formula = '"' + $ComposerServersImport.Build + '"'
	$ComposerObject.Cells( "Prop.ApiVersion" ).Formula = '"' + $ComposerServersImport.ApiVersion + '"'
	$ComposerObject.Cells( "Prop.MinVCVersion" ).Formula = '"' + $ComposerServersImport.MinVCVersion + '"'
	$ComposerObject.Cells( "Prop.MinESXVersion" ).Formula = '"' + $ComposerServersImport.MinESXVersion + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_Id" ).Formula = '"' + $ComposerServersImport.ConnectionServer_Id + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_Name" ).Formula = '"' + $ComposerServersImport.ConnectionServer_Name + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_Status" ).Formula = '"' + $ComposerServersImport.ConnectionServer_Status + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_ErrorMessage" ).Formula = '"' + $ComposerServersImport.ConnectionServer_ErrorMessage + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_ThumbprintAccepted" ).Formula = '"' + $ComposerServersImport.ConnectionServer_ThumbprintAccepted + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_CertificateHealth_Valid" ).Formula = '"' + $ComposerServersImport.ConnectionServer_CertificateHealth_Valid + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_CertificateHealth_StartTime" ).Formula = '"' + $ComposerServersImport.ConnectionServer_CertificateHealth_StartTime + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_CertificateHealth_ExpirationTime" ).Formula = '"' + $ComposerServersImport.ConnectionServer_CertificateHealth_ExpirationTime + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_CertificateHealth_InvalidReason" ).Formula = '"' + $ComposerServersImport.ConnectionServer_CertificateHealth_InvalidReason + '"'
	$ComposerObject.Cells( "Prop.ConnectionServer_CertificateHealth_ConnectionServerCertificate" ).Formula = '"' + $ComposerServersImport.ConnectionServer_CertificateHealth_ConnectionServerCertificate + '"'
}
#endregion ~~< Draw_Composer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Connection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Connection
{ `
	$ConnectionServerObject.Cells( "Prop.Id" ).Formula = '"' + $ConnectionServer.Id + '"'
	$ConnectionServerObject.Cells( "Prop.Name" ).Formula = '"' + $ConnectionServer.Name + '"'
	$ConnectionServerObject.Cells( "Prop.ServerAddress" ).Formula = '"' + $ConnectionServer.ServerAddress + '"'
	$ConnectionServerObject.Cells( "Prop.Enabled" ).Formula = '"' + $ConnectionServer.Enabled + '"'
	$ConnectionServerObject.Cells( "Prop.Tags" ).Formula = '"' + $ConnectionServer.Tags + '"'
	$ConnectionServerObject.Cells( "Prop.ExternalURL" ).Formula = '"' + $ConnectionServer.ExternalURL + '"'
	$ConnectionServerObject.Cells( "Prop.ExternalPCoIPURL" ).Formula = '"' + $ConnectionServer.ExternalPCoIPURL + '"'
	$ConnectionServerObject.Cells( "Prop.HasPCoIPGatewaySupport" ).Formula = '"' + $ConnectionServer.HasPCoIPGatewaySupport + '"'
	$ConnectionServerObject.Cells( "Prop.HasBlastGatewaySupport" ).Formula = '"' + $ConnectionServer.HasBlastGatewaySupport + '"'
	$ConnectionServerObject.Cells( "Prop.AuxillaryExternalPCoIPIPv4Address" ).Formula = '"' + $ConnectionServer.AuxillaryExternalPCoIPIPv4Address + '"'
	$ConnectionServerObject.Cells( "Prop.ExternalAppblastURL" ).Formula = '"' + $ConnectionServer.ExternalAppblastURL + '"'
	$ConnectionServerObject.Cells( "Prop.LocalConnectionServer" ).Formula = '"' + $ConnectionServer.LocalConnectionServer + '"'
	$ConnectionServerObject.Cells( "Prop.BypassTunnel" ).Formula = '"' + $ConnectionServer.BypassTunnel + '"'
	$ConnectionServerObject.Cells( "Prop.BypassPCoIPGateway" ).Formula = '"' + $ConnectionServer.BypassPCoIPGateway + '"'
	$ConnectionServerObject.Cells( "Prop.BypassAppBlastGateway" ).Formula = '"' + $ConnectionServer.BypassAppBlastGateway + '"'
	$ConnectionServerObject.Cells( "Prop.DirectHTMLABSG" ).Formula = '"' + $ConnectionServer.DirectHTMLABSG + '"'
	$ConnectionServerObject.Cells( "Prop.FullVersion" ).Formula = '"' + $ConnectionServer.FullVersion + '"'
	$ConnectionServerObject.Cells( "Prop.IpMode" ).Formula = '"' + $ConnectionServer.IpMode + '"'
	$ConnectionServerObject.Cells( "Prop.FipsModeEnabled" ).Formula = '"' + $ConnectionServer.FipsModeEnabled + '"'
	$ConnectionServerObject.Cells( "Prop.Fqhn" ).Formula = '"' + $ConnectionServer.Fqhn + '"'
	$ConnectionServerObject.Cells( "Prop.SmartCardSupport" ).Formula = '"' + $ConnectionServer.SmartCardSupport + '"'
	$ConnectionServerObject.Cells( "Prop.EnableSmartCardUserNameHint" ).Formula = '"' + $ConnectionServer.EnableSmartCardUserNameHint + '"'
	$ConnectionServerObject.Cells( "Prop.LogoffWhenRemoveSmartCard" ).Formula = '"' + $ConnectionServer.LogoffWhenRemoveSmartCard + '"'
	$ConnectionServerObject.Cells( "Prop.SmartCardSupportForAdmin" ).Formula = '"' + $ConnectionServer.SmartCardSupportForAdmin + '"'
	$ConnectionServerObject.Cells( "Prop.RsaSecureIdConfig_SecureIdEnabled" ).Formula = '"' + $ConnectionServer.RsaSecureIdConfig_SecureIdEnabled + '"'
	$ConnectionServerObject.Cells( "Prop.RsaSecureIdConfig_NameMapping" ).Formula = '"' + $ConnectionServer.RsaSecureIdConfig_NameMapping + '"'
	$ConnectionServerObject.Cells( "Prop.RsaSecureIdConfig_ClearNodeSecret" ).Formula = '"' + $ConnectionServer.RsaSecureIdConfig_ClearNodeSecret + '"'
	$ConnectionServerObject.Cells( "Prop.RsaSecureIdConfig_SecurityFileData" ).Formula = '"' + $ConnectionServer.RsaSecureIdConfig_SecurityFileData + '"'
	$ConnectionServerObject.Cells( "Prop.RsaSecureIdConfig_SecurityFileUploaded" ).Formula = '"' + $ConnectionServer.RsaSecureIdConfig_SecurityFileUploaded + '"'
	$ConnectionServerObject.Cells( "Prop.RadiusConfig_RadiusEnabled" ).Formula = '"' + $ConnectionServer.RadiusConfig_RadiusEnabled + '"'
	$ConnectionServerObject.Cells( "Prop.RadiusConfig_RadiusAuthenticator" ).Formula = '"' + $ConnectionServer.RadiusConfig_RadiusAuthenticator + '"'
	$ConnectionServerObject.Cells( "Prop.RadiusConfig_RadiusNameMapping" ).Formula = '"' + $ConnectionServer.RadiusConfig_RadiusNameMapping + '"'
	$ConnectionServerObject.Cells( "Prop.RadiusConfig_RadiusSSO" ).Formula = '"' + $ConnectionServer.RadiusConfig_RadiusSSO + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_SamlSupport" ).Formula = '"' + $ConnectionServer.SamlConfig_SamlSupport + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_SamlAuthenticator_Id" ).Formula = '"' + $ConnectionServer.SamlConfig_SamlAuthenticator_Id + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_SamlAuthenticators_Id" ).Formula = '"' + $ConnectionServer.SamlConfig_SamlAuthenticators_Id + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_WorkspaceOneData_WorkspaceOneModeEnabled" ).Formula = '"' + $ConnectionServer.SamlConfig_WorkspaceOneData_WorkspaceOneModeEnabled + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_WorkspaceOneData_WorkspaceOneHostName" ).Formula = '"' + $ConnectionServer.SamlConfig_WorkspaceOneData_WorkspaceOneHostName + '"'
	$ConnectionServerObject.Cells( "Prop.SamlConfig_WorkspaceOneData_WorkspaceOneBlockOldClients" ).Formula = '"' + $ConnectionServer.SamlConfig_WorkspaceOneData_WorkspaceOneBlockOldClients + '"'
	$ConnectionServerObject.Cells( "Prop.UnauthenticatedAccessConfig_Enabled" ).Formula = '"' + $ConnectionServer.UnauthenticatedAccessConfig_Enabled + '"'
	$ConnectionServerObject.Cells( "Prop.UnauthenticatedAccessConfig_DefaultUser" ).Formula = '"' + $ConnectionServer.UnauthenticatedAccessConfig_DefaultUser + '"'
	$ConnectionServerObject.Cells( "Prop.UnauthenticatedAccessConfig_UserIdleTimeout" ).Formula = '"' + $ConnectionServer.UnauthenticatedAccessConfig_UserIdleTimeout + '"'
	$ConnectionServerObject.Cells( "Prop.UnauthenticatedAccessConfig_ClientPuzzleDifficulty" ).Formula = '"' + $ConnectionServer.UnauthenticatedAccessConfig_ClientPuzzleDifficulty + '"'
	$ConnectionServerObject.Cells( "Prop.UnauthenticatedAccessConfig_BlockUnsupportedClients" ).Formula = '"' + $ConnectionServer.UnauthenticatedAccessConfig_BlockUnsupportedClients + '"'
	$ConnectionServerObject.Cells( "Prop.LdapBackupFrequencyTime" ).Formula = '"' + $ConnectionServer.LdapBackupFrequencyTime + '"'
	$ConnectionServerObject.Cells( "Prop.LdapBackupMaxNumber" ).Formula = '"' + $ConnectionServer.LdapBackupMaxNumber + '"'
	$ConnectionServerObject.Cells( "Prop.LdapBackupFolder" ).Formula = '"' + $ConnectionServer.LdapBackupFolder + '"'
	$ConnectionServerObject.Cells( "Prop.LastLdapBackupTime" ).Formula = '"' + $ConnectionServer.LastLdapBackupTime + '"'
	$ConnectionServerObject.Cells( "Prop.LastLdapBackupStatus" ).Formula = '"' + $ConnectionServer.LastLdapBackupStatus + '"'
	$ConnectionServerObject.Cells( "Prop.IsBackupInProgress" ).Formula = '"' + $ConnectionServer.IsBackupInProgress + '"'
	$ConnectionServerObject.Cells( "Prop.LdapBackupTimeOffset" ).Formula = '"' + $ConnectionServer.LdapBackupTimeOffset + '"'
	$ConnectionServerObject.Cells( "Prop.SecurityServerPairing" ).Formula = '"' + $ConnectionServer.SecurityServerPairing + '"'
	$ConnectionServerObject.Cells( "Prop.MessageSecurity_MessageSecurityEnhancedModeSupported" ).Formula = '"' + $ConnectionServer.MessageSecurity_MessageSecurityEnhancedModeSupported + '"'
	$ConnectionServerObject.Cells( "Prop.MessageSecurity_RouterSslThumbprints" ).Formula = '"' + $ConnectionServer.MessageSecurity_RouterSslThumbprints + '"'
	$ConnectionServerObject.Cells( "Prop.MessageSecurity_MsgSecurityPublicKey" ).Formula = '"' + $ConnectionServer.MessageSecurity_MsgSecurityPublicKey + '"'
	$ConnectionServerObject.Cells( "Prop.Status" ).Formula = '"' + $ConnectionServer.Status + '"'
	$ConnectionServerObject.Cells( "Prop.Version" ).Formula = '"' + $ConnectionServer.Version + '"'
	$ConnectionServerObject.Cells( "Prop.Build" ).Formula = '"' + $ConnectionServer.Build + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumConnections" ).Formula = '"' + $ConnectionServer.ConnectionData_NumConnections + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumConnectionsHigh" ).Formula = '"' + $ConnectionServer.ConnectionData_NumConnectionsHigh + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumViewComposerConnections" ).Formula = '"' + $ConnectionServer.ConnectionData_NumViewComposerConnections + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumViewComposerConnectionsHigh" ).Formula = '"' + $ConnectionServer.ConnectionData_NumViewComposerConnectionsHigh + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumTunneledSessions" ).Formula = '"' + $ConnectionServer.ConnectionData_NumTunneledSessions + '"'
	$ConnectionServerObject.Cells( "Prop.ConnectionData_NumPSGSessions" ).Formula = '"' + $ConnectionServer.ConnectionData_NumPSGSessions + '"'
	$ConnectionServerObject.Cells( "Prop.DefaultCertificate" ).Formula = '"' + $ConnectionServer.DefaultCertificate + '"'
	$ConnectionServerObject.Cells( "Prop.CertificateHealth_Valid" ).Formula = '"' + $ConnectionServer.CertificateHealth_Valid + '"'
	$ConnectionServerObject.Cells( "Prop.CertificateHealth_StartTime" ).Formula = '"' + $ConnectionServer.CertificateHealth_StartTime + '"'
	$ConnectionServerObject.Cells( "Prop.CertificateHealth_ExpirationTime" ).Formula = '"' + $ConnectionServer.CertificateHealth_ExpirationTime + '"'
	$ConnectionServerObject.Cells( "Prop.CertificateHealth_InvalidReason" ).Formula = '"' + $ConnectionServer.CertificateHealth_InvalidReason + '"'
	$ConnectionServerObject.Cells( "Prop.CertificateHealth_ConnectionServerCertificate" ).Formula = '"' + $ConnectionServer.CertificateHealth_ConnectionServerCertificate + '"'
}
#endregion ~~< Draw_Connection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Pool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Pool
{ `
	$PoolObject.Cells( "Prop.Id" ).Formula = '"' + $Pool.Id + '"'
	$PoolObject.Cells( "Prop.Name" ).Formula = '"' + $Pool.Name + '"'
	$PoolObject.Cells( "Prop.DisplayName" ).Formula = '"' + $Pool.DisplayName + '"'
	$PoolObject.Cells( "Prop.Enabled" ).Formula = '"' + $Pool.Enabled + '"'
	$PoolObject.Cells( "Prop.Deleting" ).Formula = '"' + $Pool.Deleting + '"'
	$PoolObject.Cells( "Prop.Type" ).Formula = '"' + $Pool.Type + '"'
	$PoolObject.Cells( "Prop.Source" ).Formula = '"' + $Pool.Source + '"'
	$PoolObject.Cells( "Prop.UserAssignment" ).Formula = '"' + $Pool.UserAssignment + '"'
	$PoolObject.Cells( "Prop.AccessGroup_Id" ).Formula = '"' + $Pool.AccessGroup_Id + '"'
	$PoolObject.Cells( "Prop.GlobalEntitlement" ).Formula = '"' + $Pool.GlobalEntitlement + '"'
	$PoolObject.Cells( "Prop.VirtualCenter_Id" ).Formula = '"' + $Pool.VirtualCenter_Id + '"'
	$PoolObject.Cells( "Prop.ProvisioningEnabled" ).Formula = '"' + $Pool.ProvisioningEnabled + '"'
	$PoolObject.Cells( "Prop.NumMachines" ).Formula = '"' + $Pool.NumMachines + '"'
	$PoolObject.Cells( "Prop.NumSessions" ).Formula = '"' + $Pool.NumSessions + '"'
	$PoolObject.Cells( "Prop.Farm_Id" ).Formula = '"' + $Pool.Farm_Id + '"'
	$PoolObject.Cells( "Prop.SupportedDomains" ).Formula = '"' + $Pool.SupportedDomains + '"'
	$PoolObject.Cells( "Prop.LastProvisioningError" ).Formula = '"' + $Pool.LastProvisioningError + '"'
	$PoolObject.Cells( "Prop.CategoryFolderName" ).Formula = '"' + $Pool.CategoryFolderName + '"'
	$PoolObject.Cells( "Prop.EnableAppRemoting" ).Formula = '"' + $Pool.EnableAppRemoting + '"'
	$PoolObject.Cells( "Prop.ApplicationCount" ).Formula = '"' + $Pool.ApplicationCount + '"'
	$PoolObject.Cells( "Prop.SupportedSessionType" ).Formula = '"' + $Pool.SupportedSessionType + '"'
	$PoolObject.Cells( "Prop.OperatingSystem" ).Formula = '"' + $Pool.OperatingSystem + '"'
	$PoolObject.Cells( "Prop.OperatingSystemArchitecture" ).Formula = '"' + $Pool.OperatingSystemArchitecture + '"'
	$PoolObject.Cells( "Prop.EnableGRIDvGPUs" ).Formula = '"' + $Pool.EnableGRIDvGPUs + '"'
	$PoolObject.Cells( "Prop.Renderer3D" ).Formula = '"' + $Pool.Renderer3D + '"'
	$PoolObject.Cells( "Prop.AllowUsersToChooseProtocol" ).Formula = '"' + $Pool.AllowUsersToChooseProtocol + '"'
	$PoolObject.Cells( "Prop.AllowMultipleSessionsPerUser" ).Formula = '"' + $Pool.AllowMultipleSessionsPerUser + '"'
	$PoolObject.Cells( "Prop.AllowUsersToResetMachines" ).Formula = '"' + $Pool.AllowUsersToResetMachines + '"'
	$PoolObject.Cells( "Prop.DefaultDisplayProtocol" ).Formula = '"' + $Pool.DefaultDisplayProtocol + '"'
	$PoolObject.Cells( "Prop.EnableHTMLAccess" ).Formula = '"' + $Pool.EnableHTMLAccess + '"'
	$PoolObject.Cells( "Prop.EnableCollaboration" ).Formula = '"' + $Pool.EnableCollaboration + '"'
	$PoolObject.Cells( "Prop.MultipleSessionAutoClean" ).Formula = '"' + $Pool.MultipleSessionAutoClean + '"'
}
#endregion ~~< Draw_Pool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Desktop
{ `
	$DesktopObject.Cells( "Prop.Id" ).Formula = '"' + $Desktop.Id + '"'
	$DesktopObject.Cells( "Prop.GroupId" ).Formula = '"' + $Desktop.GroupId + '"'
	$DesktopObject.Cells( "Prop.Name" ).Formula = '"' + $Desktop.Name + '"'
	$DesktopObject.Cells( "Prop.AssignedUser_Id" ).Formula = '"' + $Desktop.AssignedUser_Id + '"'
	$DesktopObject.Cells( "Prop.AssignedUserName" ).Formula = '"' + $Desktop.AssignedUserName + '"'
	$DesktopObject.Cells( "Prop.Type" ).Formula = '"' + $Desktop.Type + '"'
	$DesktopObject.Cells( "Prop.Source" ).Formula = '"' + $Desktop.Source + '"'
	$DesktopObject.Cells( "Prop.UserAssignment" ).Formula = '"' + $Desktop.UserAssignment + '"'
	$DesktopObject.Cells( "Prop.SessionProtocol" ).Formula = '"' + $Desktop.SessionProtocol + '"'
	$DesktopObject.Cells( "Prop.VirtualCenter_Id" ).Formula = '"' + $Desktop.VirtualCenter_Id + '"'
	$DesktopObject.Cells( "Prop.VirtualDisks_Path" ).Formula = '"' + $Desktop.VirtualDisks_Path + '"'
	$DesktopObject.Cells( "Prop.VirtualDisks_DatastorePath" ).Formula = '"' + $Desktop.VirtualDisks_DatastorePath + '"'
	$DesktopObject.Cells( "Prop.VirtualDisks_CapacityMB" ).Formula = '"' + $Desktop.VirtualDisks_CapacityMB + '"'
	$DesktopObject.Cells( "Prop.PersistentDisks" ).Formula = '"' + $Desktop.PersistentDisks + '"'
	$DesktopObject.Cells( "Prop.LastMaintenanceTime" ).Formula = '"' + $Desktop.LastMaintenanceTime + '"'
	$DesktopObject.Cells( "Prop.Operation" ).Formula = '"' + $Desktop.Operation + '"'
	$DesktopObject.Cells( "Prop.OperationState" ).Formula = '"' + $Desktop.OperationState + '"'
	$DesktopObject.Cells( "Prop.AutoRefreshLogOffSetting" ).Formula = '"' + $Desktop.AutoRefreshLogOffSetting + '"'
	$DesktopObject.Cells( "Prop.InHoldCustomization" ).Formula = '"' + $Desktop.InHoldCustomization + '"'
	$DesktopObject.Cells( "Prop.MissingInVCenter" ).Formula = '"' + $Desktop.MissingInVCenter + '"'
	$DesktopObject.Cells( "Prop.CreateTime" ).Formula = '"' + $Desktop.CreateTime + '"'
	$DesktopObject.Cells( "Prop.CloneErrorMessage" ).Formula = '"' + $Desktop.CloneErrorMessage + '"'
	$DesktopObject.Cells( "Prop.CloneErrorTime" ).Formula = '"' + $Desktop.CloneErrorTime + '"'
	$DesktopObject.Cells( "Prop.BaseImagePath" ).Formula = '"' + $Desktop.BaseImagePath + '"'
	$DesktopObject.Cells( "Prop.BaseImageSnapshotPath" ).Formula = '"' + $Desktop.BaseImageSnapshotPath + '"'
	$DesktopObject.Cells( "Prop.PendingBaseImagePath" ).Formula = '"' + $Desktop.PendingBaseImagePath + '"'
	$DesktopObject.Cells( "Prop.PendingBaseImageSnapshotPath" ).Formula = '"' + $Desktop.PendingBaseImageSnapshotPath + '"'
	$DesktopObject.Cells( "Prop.PairingState" ).Formula = '"' + $Desktop.PairingState + '"'
	$DesktopObject.Cells( "Prop.ConfiguredByBroker" ).Formula = '"' + $Desktop.ConfiguredByBroker + '"'
	$DesktopObject.Cells( "Prop.AttemptedTheftByBroker" ).Formula = '"' + $Desktop.AttemptedTheftByBroker + '"'
	$DesktopObject.Cells( "Prop.MachinePowerState" ).Formula = '"' + $Desktop.MachinePowerState + '"'
	$DesktopObject.Cells( "Prop.IpV4" ).Formula = '"' + $Desktop.IpV4 + '"'
	$DesktopObject.Cells( "Prop.IpV6" ).Formula = '"' + $Desktop.IpV6 + '"'
	$DesktopObject.Cells( "Prop.AgentId" ).Formula = '"' + $Desktop.AgentId + '"'
	$DesktopObject.Cells( "Prop.DnsName" ).Formula = '"' + $Desktop.DnsName + '"'
	$DesktopObject.Cells( "Prop.User_Id" ).Formula = '"' + $Desktop.User_Id + '"'
	$DesktopObject.Cells( "Prop.AccessGroup_Id" ).Formula = '"' + $Desktop.AccessGroup_Id + '"'
	$DesktopObject.Cells( "Prop.Desktop_Id" ).Formula = '"' + $Desktop.Desktop_Id + '"'
	$DesktopObject.Cells( "Prop.DesktopName" ).Formula = '"' + $Desktop.DesktopName + '"'
	$DesktopObject.Cells( "Prop.Session_Id" ).Formula = '"' + $Desktop.Session_Id + '"'
	$DesktopObject.Cells( "Prop.BasicState" ).Formula = '"' + $Desktop.BasicState + '"'
	$DesktopObject.Cells( "Prop.Base_Type" ).Formula = '"' + $Desktop.Base_Type + '"'
	$DesktopObject.Cells( "Prop.OperatingSystem" ).Formula = '"' + $Desktop.OperatingSystem + '"'
	$DesktopObject.Cells( "Prop.OperatingSystemArchitecture" ).Formula = '"' + $Desktop.OperatingSystemArchitecture + '"'
	$DesktopObject.Cells( "Prop.AgentVersion" ).Formula = '"' + $Desktop.AgentVersion + '"'
	$DesktopObject.Cells( "Prop.AgentBuildNumber" ).Formula = '"' + $Desktop.AgentBuildNumber + '"'
	$DesktopObject.Cells( "Prop.RemoteExperienceAgentVersion" ).Formula = '"' + $Desktop.RemoteExperienceAgentVersion + '"'
	$DesktopObject.Cells( "Prop.RemoteExperienceAgentBuildNumber" ).Formula = '"' + $Desktop.RemoteExperienceAgentBuildNumber + '"'
	$DesktopObject.Cells( "Prop.UserName" ).Formula = '"' + $Desktop.UserName + '"'
	$DesktopObject.Cells( "Prop.MessageSecurityMode" ).Formula = '"' + $Desktop.MessageSecurityMode + '"'
	$DesktopObject.Cells( "Prop.MessageSecurityEnhancedModeSupported" ).Formula = '"' + $Desktop.MessageSecurityEnhancedModeSupported + '"'
	$DesktopObject.Cells( "Prop.HostName" ).Formula = '"' + $Desktop.HostName + '"'
	$DesktopObject.Cells( "Prop.DatastorePaths" ).Formula = '"' + $Desktop.DatastorePaths + '"'
}
#endregion ~~< Draw_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_RDSH >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_RDSH
{ `
	$RDSHObject.Cells( "Prop.Id" ).Formula = '"' + $RDSServer.Id + '"'
	$RDSHObject.Cells( "Prop.Name" ).Formula = '"' + $RDSServer.Name + '"'
	$RDSHObject.Cells( "Prop.Description" ).Formula = '"' + $RDSServer.Description + '"'
	$RDSHObject.Cells( "Prop.Farm_Id" ).Formula = '"' + $RDSServer.Farm_Id + '"'
	$RDSHObject.Cells( "Prop.Desktop_Id" ).Formula = '"' + $RDSServer.Desktop_Id + '"'
	$RDSHObject.Cells( "Prop.AccessGroup_Id" ).Formula = '"' + $RDSServer.AccessGroup_Id + '"'
	$RDSHObject.Cells( "Prop.MessageSecurityMode" ).Formula = '"' + $RDSServer.MessageSecurityMode + '"'
	$RDSHObject.Cells( "Prop.MessageSecurityEnhancedModeSupported" ).Formula = '"' + $RDSServer.MessageSecurityEnhancedModeSupported + '"'
	$RDSHObject.Cells( "Prop.DnsName" ).Formula = '"' + $RDSServer.DnsName + '"'
	$RDSHObject.Cells( "Prop.OperatingSystem" ).Formula = '"' + $RDSServer.OperatingSystem + '"'
	$RDSHObject.Cells( "Prop.AgentVersion" ).Formula = '"' + $RDSServer.AgentVersion + '"'
	$RDSHObject.Cells( "Prop.AgentBuildNumber" ).Formula = '"' + $RDSServer.AgentBuildNumber + '"'
	$RDSHObject.Cells( "Prop.RemoteExperienceAgentVersion" ).Formula = '"' + $RDSServer.RemoteExperienceAgentVersion + '"'
	$RDSHObject.Cells( "Prop.RemoteExperienceAgentBuildNumber" ).Formula = '"' + $RDSServer.RemoteExperienceAgentBuildNumber + '"'
	$RDSHObject.Cells( "Prop.SessionSettings_MaxSessionsType" ).Formula = '"' + $RDSServer.SessionSettings_MaxSessionsType + '"'
	$RDSHObject.Cells( "Prop.SessionSettings_MaxSessionsSetByAdmin" ).Formula = '"' + $RDSServer.SessionSettings_MaxSessionsSetByAdmin + '"'
	$RDSHObject.Cells( "Prop.Agent_MaxSessionsType" ).Formula = '"' + $RDSServer.Agent_MaxSessionsType + '"'
	$RDSHObject.Cells( "Prop.Agent_MaxSessionsSetByAdmin" ).Formula = '"' + $RDSServer.Agent_MaxSessionsSetByAdmin + '"'
	$RDSHObject.Cells( "Prop.Enabled" ).Formula = '"' + $RDSServer.Enabled + '"'
	$RDSHObject.Cells( "Prop.Status" ).Formula = '"' + $RDSServer.Status + '"'
	$RDSHObject.Cells( "Prop.SessionCount" ).Formula = '"' + $RDSServer.SessionCount + '"'
	$RDSHObject.Cells( "Prop.LoadPreference" ).Formula = '"' + $RDSServer.LoadPreference + '"'
	$RDSHObject.Cells( "Prop.LoadIndex" ).Formula = '"' + $RDSServer.LoadIndex + '"'
	$RDSHObject.Cells( "Prop.Operation" ).Formula = '"' + $RDSServer.Operation + '"'
	$RDSHObject.Cells( "Prop.OperationState" ).Formula = '"' + $RDSServer.OperationState + '"'
	$RDSHObject.Cells( "Prop.LogOffSetting" ).Formula = '"' + $RDSServer.LogOffSetting + '"'
	$RDSHObject.Cells( "Prop.BaseImagePath" ).Formula = '"' + $RDSServer.BaseImagePath + '"'
	$RDSHObject.Cells( "Prop.BaseImageSnapshotPath" ).Formula = '"' + $RDSServer.BaseImageSnapshotPath + '"'
	$RDSHObject.Cells( "Prop.PendingBaseImagePath" ).Formula = '"' + $RDSServer.PendingBaseImagePath + '"'
	$RDSHObject.Cells( "Prop.PendingBaseImageSnapshotPath" ).Formula = '"' + $RDSServer.PendingBaseImageSnapshotPath + '"'
	$RDSHObject.Cells( "Prop.FarmName" ).Formula = '"' + $RDSServer.FarmName + '"'
	$RDSHObject.Cells( "Prop.DesktopName" ).Formula = '"' + $RDSServer.DesktopName + '"'
	$RDSHObject.Cells( "Prop.FarmType" ).Formula = '"' + $RDSServer.FarmType + '"'
	$RDSHObject.Cells( "Prop.MachinePowerState" ).Formula = '"' + $RDSServer.MachinePowerState + '"'
	$RDSHObject.Cells( "Prop.IpV4" ).Formula = '"' + $RDSServer.IpV4 + '"'
	$RDSHObject.Cells( "Prop.IpV6" ).Formula = '"' + $RDSServer.IpV6 + '"'
	$RDSHObject.Cells( "Prop.AgentId" ).Formula = '"' + $RDSServer.AgentId + '"'
}
#endregion ~~< Draw_RDSH >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Farm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Farm
{ `
	$FarmObject.Cells( "Prop.Id" ).Formula = '"' + $Farm.Id + '"'
	$FarmObject.Cells( "Prop.Name" ).Formula = '"' + $Farm.Name + '"'
	$FarmObject.Cells( "Prop.Type" ).Formula = '"' + $Farm.Type + '"'
	$FarmObject.Cells( "Prop.Health" ).Formula = '"' + $Farm.Health + '"'
	$FarmObject.Cells( "Prop.Source" ).Formula = '"' + $Farm.Source + '"'
	$FarmObject.Cells( "Prop.AccessGroup_Id" ).Formula = '"' + $Farm.AccessGroup_Id + '"'
	$FarmObject.Cells( "Prop.RdsServer_Id" ).Formula = '"' + $Farm.RdsServer_Id + '"'
	$FarmObject.Cells( "Prop.RdsServer_Name" ).Formula = '"' + $Farm.RdsServer_Name + '"'
	$FarmObject.Cells( "Prop.RdsServer_OperatingSystem" ).Formula = '"' + $Farm.RdsServer_OperatingSystem + '"'
	$FarmObject.Cells( "Prop.RdsServer_AgentVersion" ).Formula = '"' + $Farm.RdsServer_AgentVersion + '"'
	$FarmObject.Cells( "Prop.RdsServer_AgentBuildNumber" ).Formula = '"' + $Farm.RdsServer_AgentBuildNumber + '"'
	$FarmObject.Cells( "Prop.RdsServer_Status" ).Formula = '"' + $Farm.RdsServer_Status + '"'
	$FarmObject.Cells( "Prop.RdsServer_Health" ).Formula = '"' + $Farm.RdsServer_Health + '"'
	$FarmObject.Cells( "Prop.RdsServer_Available" ).Formula = '"' + $Farm.RdsServer_Available + '"'
	$FarmObject.Cells( "Prop.RdsServer_MissingApplications" ).Formula = '"' + $Farm.RdsServer_MissingApplications + '"'
	$FarmObject.Cells( "Prop.RdsServer_LoadPreference" ).Formula = '"' + $Farm.RdsServer_LoadPreference + '"'
	$FarmObject.Cells( "Prop.RdsServer_LoadIndex" ).Formula = '"' + $Farm.RdsServer_LoadIndex + '"'
	$FarmObject.Cells( "Prop.RdsServer_SessionSettings_MaxSessionsType" ).Formula = '"' + $Farm.RdsServer_SessionSettings_MaxSessionsType + '"'
	$FarmObject.Cells( "Prop.RdsServer_SessionSettings_MaxSessionsSetByAdmin" ).Formula = '"' + $Farm.RdsServer_SessionSettings_MaxSessionsSetByAdmin + '"'
	$FarmObject.Cells( "Prop.NumApplications" ).Formula = '"' + $Farm.NumApplications + '"'
	$FarmObject.Cells( "Prop.DisplayName" ).Formula = '"' + $Farm.DisplayName + '"'
	$FarmObject.Cells( "Prop.Description" ).Formula = '"' + $Farm.Description + '"'
	$FarmObject.Cells( "Prop.AccessGroupName" ).Formula = '"' + $Farm.AccessGroupName + '"'
	$FarmObject.Cells( "Prop.Enabled" ).Formula = '"' + $Farm.Enabled + '"'
	$FarmObject.Cells( "Prop.ProvisioningEnabled" ).Formula = '"' + $Farm.ProvisioningEnabled + '"'
	$FarmObject.Cells( "Prop.Deleting" ).Formula = '"' + $Farm.Deleting + '"'
	$FarmObject.Cells( "Prop.Desktop_Id" ).Formula = '"' + $Farm.Desktop_Id + '"'
	$FarmObject.Cells( "Prop.DesktopName" ).Formula = '"' + $Farm.DesktopName + '"'
	$FarmObject.Cells( "Prop.RdsServerCount" ).Formula = '"' + $Farm.RdsServerCount + '"'
}
#endregion ~~< Draw_Farm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Application
{ `
	$ApplicationObject.Cells( "Prop.Id" ).Formula = '"' + $Application.Id + '"'
	$ApplicationObject.Cells( "Prop.Name" ).Formula = '"' + $Application.Name + '"'
	$ApplicationObject.Cells( "Prop.DisplayName" ).Formula = '"' + $Application.DisplayName + '"'
	$ApplicationObject.Cells( "Prop.Description" ).Formula = '"' + $Application.Description + '"'
	$ApplicationObject.Cells( "Prop.Enabled" ).Formula = '"' + $Application.Enabled + '"'
	$ApplicationObject.Cells( "Prop.GlobalApplicationEntitlement" ).Formula = '"' + $Application.GlobalApplicationEntitlement + '"'
	$ApplicationObject.Cells( "Prop.EnableAntiAffinityRules" ).Formula = '"' + $Application.EnableAntiAffinityRules + '"'
	$ApplicationObject.Cells( "Prop.AntiAffinityPatterns" ).Formula = '"' + $Application.AntiAffinityPatterns + '"'
	$ApplicationObject.Cells( "Prop.AntiAffinityCount" ).Formula = '"' + $Application.AntiAffinityCount + '"'
	$ApplicationObject.Cells( "Prop.EnablePreLaunch" ).Formula = '"' + $Application.EnablePreLaunch + '"'
	$ApplicationObject.Cells( "Prop.ConnectionServerRestrictions" ).Formula = '"' + $Application.ConnectionServerRestrictions + '"'
	$ApplicationObject.Cells( "Prop.CategoryFolderName" ).Formula = '"' + $Application.CategoryFolderName + '"'
	$ApplicationObject.Cells( "Prop.ClientRestrictions" ).Formula = '"' + $Application.ClientRestrictions + '"'
	$ApplicationObject.Cells( "Prop.ShortcutLocations" ).Formula = '"' + $Application.ShortcutLocations + '"'
	$ApplicationObject.Cells( "Prop.MultiSessionMode" ).Formula = '"' + $Application.MultiSessionMode + '"'
	$ApplicationObject.Cells( "Prop.MaxMultiSessions" ).Formula = '"' + $Application.MaxMultiSessions + '"'
	$ApplicationObject.Cells( "Prop.ExecutablePath" ).Formula = '"' + $Application.ExecutablePath + '"'
	$ApplicationObject.Cells( "Prop.Version" ).Formula = '"' + $Application.Version + '"'
	$ApplicationObject.Cells( "Prop.Publisher" ).Formula = '"' + $Application.Publisher + '"'
	$ApplicationObject.Cells( "Prop.StartFolder" ).Formula = '"' + $Application.StartFolder + '"'
	$ApplicationObject.Cells( "Prop.Args" ).Formula = '"' + $Application.Args + '"'
	$ApplicationObject.Cells( "Prop.Farm_Id" ).Formula = '"' + $Application.Farm_Id + '"'
	$ApplicationObject.Cells( "Prop.Desktop" ).Formula = '"' + $Application.Desktop + '"'
	$ApplicationObject.Cells( "Prop.FileTypes_FileType" ).Formula = '"' + $Application.FileTypes_FileType + '"'
	$ApplicationObject.Cells( "Prop.FileTypes_Description" ).Formula = '"' + $Application.FileTypes_Description + '"'
	$ApplicationObject.Cells( "Prop.AutoUpdateFileTypes" ).Formula = '"' + $Application.AutoUpdateFileTypes + '"'
	$ApplicationObject.Cells( "Prop.OtherFileTypes" ).Formula = '"' + $Application.OtherFileTypes + '"'
	$ApplicationObject.Cells( "Prop.AutoUpdateOtherFileTypes" ).Formula = '"' + $Application.AutoUpdateOtherFileTypes + '"'
	$ApplicationObject.Cells( "Prop.Icons_Id" ).Formula = '"' + $Application.Icons_Id + '"'
}
#endregion ~~< Draw_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Gateway >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Gateway
{ `
	$GatewayObject.Cells( "Prop.Id" ).Formula = '"' + $Gateway.Id + '"'
	$GatewayObject.Cells( "Prop.Name" ).Formula = '"' + $Gateway.Name + '"'
	$GatewayObject.Cells( "Prop.Address" ).Formula = '"' + $Gateway.Address + '"'
	$GatewayObject.Cells( "Prop.GatewayZoneInternal" ).Formula = '"' + $Gateway.GatewayZoneInternal + '"'
	$GatewayObject.Cells( "Prop.Version" ).Formula = '"' + $Gateway.Version + '"'
	$GatewayObject.Cells( "Prop.Type" ).Formula = '"' + $Gateway.Type + '"'
	$GatewayObject.Cells( "Prop.ConnectionData_NumActiveConnections" ).Formula = '"' + $Gateway.ConnectionData_NumActiveConnections + '"'
	$GatewayObject.Cells( "Prop.ConnectionData_NumPcoipConnections" ).Formula = '"' + $Gateway.ConnectionData_NumPcoipConnections + '"'
	$GatewayObject.Cells( "Prop.ConnectionData_NumBlastConnections" ).Formula = '"' + $Gateway.ConnectionData_NumBlastConnections + '"'
	$GatewayObject.Cells( "Prop.GatewayStatusActive" ).Formula = '"' + $Gateway.GatewayStatusActive + '"'
	$GatewayObject.Cells( "Prop.GatewayStatusStale" ).Formula = '"' + $Gateway.GatewayStatusStale + '"'
	$GatewayObject.Cells( "Prop.GatewayContacted" ).Formula = '"' + $Gateway.GatewayContacted + '"'
}
#endregion ~~< Draw_Gateway >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Draw Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CSV Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function CSV_In_Out
{ `
	$global:DrawCsvFolder = $DrawCsvBrowse.SelectedPath

	# VirtualCenter Server
	$global:VirtualCenterExportFile = "$DrawCsvFolder\$ConnServ-VirtualCenterExport.csv"
	if ( Test-Path $VirtualCenterExportFile ) `
	{ `
		$global:VirtualCenterImport = Import-Csv $VirtualCenterExportFile
	}
	# Composer Servers
		$global:ComposerServersExportFile = "$DrawCsvFolder\$ConnServ-ComposerServersExport.csv"
	if ( Test-Path $ComposerServersExportFile ) `
	{ `
		$global:ComposerServersImport = Import-Csv $ComposerServersExportFile
	}
	# Connection Servers
		$global:ConnectionServersExportFile = "$DrawCsvFolder\$ConnServ-ConnectionServersExport.csv"
	if ( Test-Path $ConnectionServersExportFile ) `
	{ `
		$global:ConnectionServersImport = Import-Csv $ConnectionServersExportFile
	}
	# Pools
		$global:PoolsExportFile = "$DrawCsvFolder\$ConnServ-PoolsExport.csv"
	if ( Test-Path $PoolsExportFile ) `
	{ `
		$global:PoolsImport = Import-Csv $PoolsExportFile
	}
	# Desktops
	$global:DesktopsExportFile = "$DrawCsvFolder\$ConnServ-DesktopsExport.csv"
	if ( Test-Path $DesktopsExportFile ) `
	{ `
		$global:DesktopsImport = Import-Csv $DesktopsExportFile
	}
	# RDS Servers
	$global:RDSServersExportFile = "$DrawCsvFolder\$ConnServ-RDSServersExport.csv"
	if ( Test-Path $RDSServersExportFile ) `
	{ `
		$global:RDSServersImport = Import-Csv $RDSServersExportFile
	}
	# Farms
	$global:FarmsExportFile = "$DrawCsvFolder\$ConnServ-FarmsExport.csv"
	if ( Test-Path $FarmsExportFile ) `
	{ `
		$global:FarmsImport = Import-Csv $FarmsExportFile
	}
	# Applications
	$global:ApplicationsExportFile = "$DrawCsvFolder\$ConnServ-ApplicationsExport.csv"
	if ( Test-Path $ApplicationsExportFile ) `
	{ `
		$global:ApplicationsImport = Import-Csv $ApplicationsExportFile
	}
	# Gateways
	$global:GatewaysExportFile = "$DrawCsvFolder\$ConnServ-GatewaysExport.csv"
	if ( Test-Path $GatewaysExportFile ) `
	{ `
		$global:GatewaysImport = Import-Csv $GatewaysExportFile
	}
}
#endregion ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< CSV Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Shapes Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio_Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Visio_Shapes
{ `
	$stnPath = [System.Environment]::GetFolderPath( 'MyDocuments' ) + "\My Shapes"
	$stnObj = $AppVisio.Documents.Add( $stnPath + $shpFile)
	# VirtualCenter Object
	$global:VirtualCenterObj = $stnObj.Masters.Item( "VirtualCenter" )
	# Composer Object
	$global:ComposerObj = $stnObj.Masters.Item( "Composer" )
	# Connection Object
	$global:ConnectionObj = $stnObj.Masters.Item( "Connection" )
	# Pools Object
	$global:PoolObj = $stnObj.Masters.Item( "Pool" )
	# Windows Object
	$global:WindowsObj = $stnObj.Masters.Item( "Windows" )
	# Linux Object
	$global:LinuxObj = $stnObj.Masters.Item( "Linux" )
	# RDSH Object
	$global:RDSHObj = $stnObj.Masters.Item( "RDSH" )
	# Farm Object
	$global:FarmObj = $stnObj.Masters.Item( "Farm" )
	# Application Object
	$global:ApplicationObj = $stnObj.Masters.Item( "Application" )
	# Gateway Object
	$global:GatewayObj = $stnObj.Masters.Item( "Universal Access Gateway" )
}
#endregion ~~< Visio_Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Shapes Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Create_Visio_Base >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Create_Visio_Base
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Visio default page and loading shapes." -ForegroundColor Green
	$global:ConnServ = $ConnServTextBox.Text
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$DocObj = $docsObj.Add( "" )
	$DocObj.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< Create_Visio_Base >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Infrastructure >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Infrastructure
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Infrastructure Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Horizon Infrastructure Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Infrastructure"
	$DocsObj.Pages( 'Infrastructure' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Infrastructure' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$ComposerServersTotal = ( ( $ComposerServersImport ).ServerName ).Count
	$ConnectionServersNumber = 0
	$ConnectionServersTotal = ( ( $ConnectionServersImport ).Name ).Count
	$GatewaysNumber = 0
	$GatewaysTotal = ( ( $GatewaysImport ).Name ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $ComposerServersTotal + ( ( $VirtualCenterImport ).ServerName ).Count + $ConnectionServersTotal + $GatewaysTotal
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	if ( $null -ne $ComposerServersImport ) `
	{ `
		$ComposerObject = Add-VisioObjectComposer $ComposerObj $ComposerServersImport
		Draw_Composer
		$ObjectNumber++
		$Infrastructure_Complete.Forecolor = "Blue"
		$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $Infrastructure_Complete )
	
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Composer Server object -" $ComposerServersImport.ServerName
		}

		ForEach ( $VirtualCenter in ( $VirtualCenterImport | Sort-Object ServerName ) ) `
		{ `
			$x = -2.00
			$y += -1.50
			$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
			Draw_VirtualCenter
			$ObjectNumber++
			$Infrastructure_Complete.Forecolor = "Blue"
			$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $Infrastructure_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
			}

			Connect-VisioObject $VirtualCenterObject $ComposerObject
					
			ForEach ( $ConnectionServer in ( $ConnectionServersImport | Sort-Object Name -Descending | Where-Object { $VirtualCenter.ConnectionServerData_Id.contains( $_.Id ) } ) ) `
			{ `
				$x += -2.00
				$y = -1.50
				$ConnectionServerObject = Add-VisioObjectConnection $ConnectionObj $ConnectionServer
				Draw_Connection
				$ObjectNumber++
				$Infrastructure_Complete.Forecolor = "Blue"
				$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Infrastructure_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ConnectionServersNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Connection Server object " $ConnectionServersNumber " of " $ConnectionServersTotal " - " $ConnectionServer.Name
				}
				Connect-VisioObject $ConnectionServerObject $VirtualCenterObject
				$ConnectionServerObject = $VirtualCenterObject
			}
			$x = ( $x -2.00 )
			
			ForEach ( $Gateway in ( $GatewaysImport | Sort-Object Name ) ) `
			{ `
				$y += 2.00
				$GatewayObject = Add-VisioObjectGateway $GatewayObj $Gateway
				Draw_Gateway
				$ObjectNumber++
				$Infrastructure_Complete.Forecolor = "Blue"
				$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Infrastructure_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$GatewaysNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Gateway object " $GatewaysNumber " of " $GatewaysTotal " - " $Gateway.Name
				}
			}
		}
	}
	else `
	{ `
		ForEach ( $VirtualCenter in ( $VirtualCenterImport | Sort-Object ServerName ) ) `
		{ `
			$x = -2.00
			$y += -1.50
			$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
			Draw_VirtualCenter
			$ObjectNumber++
			$Infrastructure_Complete.Forecolor = "Blue"
			$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $Infrastructure_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
			}
					
			ForEach ( $ConnectionServer in ( $ConnectionServersImport | Sort-Object Name -Descending | Where-Object { $VirtualCenter.ConnectionServerData_Id.contains( $_.Id ) } ) ) `
			{ `
				$x += -2.00
				$y = -1.50
				$ConnectionServerObject = Add-VisioObjectConnection $ConnectionObj $ConnectionServer
				Draw_Connection
				$ObjectNumber++
				$Infrastructure_Complete.Forecolor = "Blue"
				$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Infrastructure_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ConnectionServersNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Connection Server object " $ConnectionServersNumber " of " $ConnectionServersTotal " - " $ConnectionServer.Name
				}
				Connect-VisioObject $ConnectionServerObject $VirtualCenterObject
				$ConnectionServerObject = $VirtualCenterObject
			}
			$x = ( $x -2.00 )
			
			ForEach ( $Gateway in ( $GatewaysImport | Sort-Object Name ) ) `
			{ `
				$y += 2.00
				$GatewayObject = Add-VisioObjectGateway $GatewayObj $Gateway
				Draw_Gateway
				$ObjectNumber++
				$Infrastructure_Complete.Forecolor = "Blue"
				$Infrastructure_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Infrastructure_Complete)
		
				if ( $debug -eq $true ) `
				{ `
					$GatewaysNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Gateway object " $GatewaysNumber " of " $GatewaysTotal " - " $Gateway.Name
				}
			}
		}
	}	
	
	# Resize to fit page
	$AppVisio.Documents.SaveAs( $SaveFile )
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< Infrastructure >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Pool_to_Desktop
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Pool to Desktop Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Pool to Desktop Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Pool to Desktop"
	$DocsObj.Pages( 'Pool to Desktop' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Pool to Desktop' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = $PoolsImport.Name.Count
	$DesktopNumber = 0
	$DesktopTotal = ( ( $DesktopsImport ).Name ).Count	
	$RDSServerNumber = 0
	$RDSServerTotal = ( $RDSServersImport | Where-Object { $_.Desktop_Id -ne "" } ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $DesktopTotal + $RDSServerTotal + ( ( $VirtualCenterImport ).ServerName ).Count
	
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
	
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$Pool_to_Desktop_Complete.Forecolor = "Blue"
		$Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}
		
		ForEach ( $Pool in ( $PoolsImport | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$Pool_to_Desktop_Complete.Forecolor = "Blue"
			$Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
			
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50
			
			ForEach ( $Desktop in ( $DesktopsImport | Sort-Object Name | Where-Object { $Pool.Id -eq ( $_.Desktop_Id ) } ) ) `
			{ `
				$x += 2.00
				if ( $Desktop.OperatingSystem.contains( "Windows" ) -eq $True ) `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $WindowsObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$Pool_to_Desktop_Complete.Forecolor = "Blue"
					$Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
					
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
				else `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $LinuxObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$Pool_to_Desktop_Complete.Forecolor = "Blue"
					$Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
					
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
			}
			
			ForEach ( $RDSServer in ( $RDSServersImport | Sort-Object Name | Where-Object { $Pool.Id -eq ( $_.Desktop_Id ) } ) ) `
			{ `
				$x += 2.00
				$RDSHObject = Add-VisioObjectRDSH $RDSHObj $RDSServer
				Draw_RDSH
				$ObjectNumber++
				$Pool_to_Desktop_Complete.Forecolor = "Blue"
				$Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Pool_to_Desktop_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$RDSServerNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing RDS Server object " $RDSServerNumber " of " $RDSServerTotal " - " $RDSServer.Name
				}
				Connect-VisioObject $PoolObject $RDSHObject
				$PoolObject = $RDSHObject 
			}
		}
	}
	
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function LC_Pool_to_Desktop
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Linked Clone Pool to Desktop Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Linked Clone Pool to Desktop Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Linked Clone Pool to Desktop"
	$DocsObj.Pages( 'Linked Clone Pool to Desktop' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Linked Clone Pool to Desktop' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$PoolsDesktops = $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" }
	$DesktopNumber = 0
	$DesktopTotal = ( ( $DesktopsImport | Sort-Object Name | Where-Object { $PoolsDesktops.Id -eq $_.Desktop_Id -and $_.Source -like "VIEW_COMPOSER" -and $PoolsDesktops.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $DesktopTotal + ( ( $VirtualCenterImport ).ServerName ).Count

	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$LC_Pool_to_Desktop_Complete.Forecolor = "Blue"
		$LC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}

		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$LC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$LC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50

			ForEach ( $Desktop in ( $DesktopsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id -and $_.Source -eq "VIEW_COMPOSER" } ) ) `
			{ `
				$x += 2.00
				if ( $Desktop.OperatingSystem.contains( "Windows" ) -eq $True ) `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $WindowsObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$LC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$LC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
				else `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $LinuxObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$LC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$LC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $LC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< LC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function LC_Pool_to_Application
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Linked Clone Pool to Application Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Linked Clone Pool to Application Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Linked Clone Pool to Application"
	$DocsObj.Pages( 'Linked Clone Pool to Application' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Linked Clone Pool to Application' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" } ).Name ).Count
	$LCPools = $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" }
	$ApplicationNumber = 0
	$ApplicationTotal = ( ( $ApplicationsImport | Sort-Object Name | Where-Object { $LCPools.Id -like $_.Desktop_Id } ).Id ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $ApplicationTotal + ( ( $VirtualCenterImport ).ServerName ).Count


	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$LC_Pool_to_Application_Complete.Forecolor = "Blue"
		$LC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}

		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "VIEW_COMPOSER" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$LC_Pool_to_Application_Complete.Forecolor = "Blue"
			$LC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50
		
			ForEach ( $Application in ( $ApplicationsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id } ) ) `
			{ `
				$x += 2.00
				$ApplicationObject = Add-VisioObjectApplication $ApplicationObj $Application
				Draw_Application
				$ObjectNumber++
				$LC_Pool_to_Application_Complete.Forecolor = "Blue"
				$LC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $LC_Pool_to_Application_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ApplicationNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Application object " $ApplicationNumber " of " $ApplicationTotal " - " $Application.Name
				}
				Connect-VisioObject $PoolObject $ApplicationObject
				$PoolObject = $ApplicationObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< LC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function FC_Pool_to_Desktop
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Full Clone Pool to Desktop Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Full Clone Pool to Desktop Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Full Clone Pool to Desktop"
	$DocsObj.Pages( 'Full Clone Pool to Desktop' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Full Clone Pool to Desktop' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$PoolsDesktops = $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" }
	$DesktopNumber = 0
	$DesktopTotal = ( ( $DesktopsImport | Sort-Object Name | Where-Object { $PoolsDesktops.Id -eq $_.Desktop_Id -and $_.Source -like "VIRTUAL_CENTER" -and $PoolsDesktops.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $DesktopTotal + ( ( $VirtualCenterImport ).ServerName ).Count

	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$FC_Pool_to_Desktop_Complete.Forecolor = "Blue"
		$FC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}

		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$FC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$FC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50
			
			ForEach ( $Desktop in ( $DesktopsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id -and $_.Source -eq "VIRTUAL_CENTER" } ) ) `
			{ `
				$x += 2.00
				if ( $Desktop.OperatingSystem.contains( "Windows" ) -eq $True ) `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $WindowsObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$FC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$FC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
				else `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $LinuxObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$FC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$FC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $FC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< FC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function FC_Pool_to_Application
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Full Clone Pool to Application Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Full Clone Pool to Application Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Full Clone Pool to Application"
	$DocsObj.Pages( 'Full Clone Pool to Application' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Full Clone Pool to Application' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" } ).Name ).Count
	$FCPools = $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" }
	$ApplicationNumber = 0
	$ApplicationTotal = ( ( $ApplicationsImport | Sort-Object Name | Where-Object { $FCPools.Id -like $_.Desktop_Id } ).Id ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $ApplicationTotal + ( ( $VirtualCenterImport ).ServerName ).Count


	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$FC_Pool_to_Application_Complete.Forecolor = "Blue"
		$FC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}

		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "VIRTUAL_CENTER" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$FC_Pool_to_Application_Complete.Forecolor = "Blue"
			$FC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50

			ForEach ( $Application in ( $ApplicationsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id } ) ) `
			{ `
				$x += 2.00
				$ApplicationObject = Add-VisioObjectApplication $ApplicationObj $Application
				Draw_Application
				$ObjectNumber++
				$FC_Pool_to_Application_Complete.Forecolor = "Blue"
				$FC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $FC_Pool_to_Application_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ApplicationNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Application object " $ApplicationNumber " of " $ApplicationTotal " - " $Application.Name
				}
				Connect-VisioObject $PoolObject $ApplicationObject
				$PoolObject = $ApplicationObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< FC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function IC_Pool_to_Desktop
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Instant Clone Pool to Desktop Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Instant Clone Pool to Desktop Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Instant Clone Pool to Desktop"
	$DocsObj.Pages( 'Instant Clone Pool to Desktop' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Instant Clone Pool to Desktop' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$PoolsDesktops = $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" }
	$DesktopNumber = 0
	$DesktopTotal = ( ( $DesktopsImport | Sort-Object Name | Where-Object { $PoolsDesktops.Id -eq $_.Desktop_Id -and $_.Source -like "INSTANT_CLONE_ENGINE" -and $PoolsDesktops.SupportedSessionType -like "*DESKTOP*" } ).Name ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $DesktopTotal + ( ( $VirtualCenterImport ).ServerName ).Count
	
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$IC_Pool_to_Desktop_Complete.Forecolor = "Blue"
		$IC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}
		
		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*DESKTOP*" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$IC_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$IC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50
			
			ForEach ( $Desktop in ( $DesktopsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id -and $_.Source -eq "INSTANT_CLONE_ENGINE" } ) ) `
			{ `
				$x += 2.00
				if ( $Desktop.OperatingSystem.contains( "Windows" ) -eq $True ) `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $WindowsObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$IC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$IC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
				else `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $LinuxObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$IC_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$IC_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $IC_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< IC_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< IC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function IC_Pool_to_Application
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Instant Clone Pool to Application Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Instant Clone Pool to Application Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Instant Clone Pool to Application"
	$DocsObj.Pages( 'Instant Clone Pool to Application' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Instant Clone Pool to Application' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" } ).Name ).Count
	$ICPools = $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" }
	$ApplicationNumber = 0
	$ApplicationTotal = ( ( $ApplicationsImport | Sort-Object Name | Where-Object { $ICPools.Id -like $_.Desktop_Id } ).Id ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $ApplicationTotal + ( ( $VirtualCenterImport ).ServerName ).Count

	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$IC_Pool_to_Application_Complete.Forecolor = "Blue"
		$IC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}
		
		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "INSTANT_CLONE_ENGINE" -and $_.Type -ne "RDS" -and $_.SupportedSessionType -like "*APPLICATION*" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$IC_Pool_to_Application_Complete.Forecolor = "Blue"
			$IC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50

			ForEach ( $Application in ( $ApplicationsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id } ) ) `
			{ `
				$x += 2.00
				$ApplicationObject = Add-VisioObjectApplication $ApplicationObj $Application
				Draw_Application
				$ObjectNumber++
				$IC_Pool_to_Application_Complete.Forecolor = "Blue"
				$IC_Pool_to_Application_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $IC_Pool_to_Application_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ApplicationNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Application object " $ApplicationNumber " of " $ApplicationTotal " - " $Application.Name
				}
				Connect-VisioObject $PoolObject $ApplicationObject
				$PoolObject = $ApplicationObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< IC_Pool_to_Application >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Unmanaged_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Unmanaged_Pool_to_Desktop
{ `
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Unmanaged Pool to Desktop Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Unmanaged Pool to Desktop Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Unmanaged Pool to Desktop"
	$DocsObj.Pages( 'Unmanaged Pool to Desktop' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Unmanaged Pool to Desktop' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Source -eq "UNMANAGED" } ).Name ).Count
	$DesktopNumber = 0
	$DesktopTotal = ( ( $DesktopsImport | Sort-Object Name | Where-Object { $PoolsImport.Id -eq $_.Desktop_Id -and $_.Source -eq "UNMANAGED" } ).Name ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $PoolTotal + $DesktopTotal + ( ( $VirtualCenterImport ).ServerName ).Count

	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Blue"
		$Unmanaged_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}
		
		ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Source -eq "UNMANAGED" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$PoolObject = Add-VisioObjectPools $PoolObj $Pool
			Draw_Pool
			$ObjectNumber++
			$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Blue"
			$Unmanaged_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$PoolNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
			}
			Connect-VisioObject $VirtualCenterObject $PoolObject
			$y += 1.50
			
			ForEach ( $Desktop in ( $DesktopsImport | Sort-Object Name | Where-Object { $Pool.Id -eq $_.Desktop_Id -and $_.Source -eq "UNMANAGED" } ) ) `
			{ `
				$x += 2.00
				if ( $Desktop.OperatingSystem.contains( "Windows" ) -eq $True ) `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $WindowsObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$Unmanaged_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
				else `
				{ `
					$DesktopObject = Add-VisioObjectDesktops $LinuxObj $Desktop
					Draw_Desktop
					$ObjectNumber++
					$Unmanaged_Pool_to_Desktop_Complete.Forecolor = "Blue"
					$Unmanaged_Pool_to_Desktop_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $Unmanaged_Pool_to_Desktop_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$DesktopNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Desktop object " $DesktopNumber " of " $DesktopTotal " - " $Desktop.Name
					}
					Connect-VisioObject $PoolObject $DesktopObject
					$PoolObject = $DesktopObject
				}
			}
		}
	}
	
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< Unmanaged_Pool_to_Desktop_Pool_to_Desktop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Farm_to_Remote_Desktop_Services >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Farm_to_Remote_Desktop_Services
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Farm to Remote Desktop Services Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Farm to Remote Desktop Services Drawing." -ForegroundColor Green
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Farm to Remote Desktop Services"
	$DocsObj.Pages( 'Farm to Remote Desktop Services' )
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item( 'Farm to Remote Desktop Services' )
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	$FarmNumber = 0
	$FarmTotal = ( ( $FarmsImport ).Name ).Count
	$PoolNumber = 0
	$PoolTotal = ( ( $PoolsImport | Where-Object { $_.Type -eq "RDS" } ).Name ).Count
	$RDSServerNumber = 0
	$RDSServerTotal = ( $RDSServersImport | Where-Object { $_.Desktop_Id -ne "" } ).Count
	$ApplicationNumber = 0
	$ApplicationTotal = ( $ApplicationsImport | Sort-Object Name | Where-Object { $FarmsImport.Id -eq $_.Farm_Id } ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $FarmTotal + $PoolTotal + $RDSServerTotal + $ApplicationTotal + ( ( $VirtualCenterImport ).ServerName ).Count


	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	ForEach ( $VirtualCenter in $VirtualCenterImport ) `
	{ `
		$VirtualCenterObject = Add-VisioObjectVirtualCenter $VirtualCenterObj $VirtualCenter
		Draw_VirtualCenter
		$ObjectNumber++
		$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
		$Farm_to_Remote_Desktop_Services_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		
		if ( $debug -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing vCenter object -" $VirtualCenter.ServerName
		}
		
		ForEach ( $Farm in ( $FarmsImport | Sort-Object Name -Descending ) ) `
		{ `
			$x = 1.50
			$y += 1.50
			$FarmObject = Add-VisioObjectFarm $FarmObj $Farm
			Draw_Farm
			$ObjectNumber++
			$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
			$Farm_to_Remote_Desktop_Services_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		
			if ( $debug -eq $true ) `
			{ `
				$FarmNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Farm object " $FarmNumber " of " $FarmTotal " - " $Farm.Name
			}
			Connect-VisioObject $VirtualCenterObject $FarmObject
			$y += 1.50
			
			ForEach ( $Pool in ( $PoolsImport | Where-Object { $_.Type -eq "RDS" -and $Farm.Id -eq $_.Farm_Id } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 3.50
				$PoolObject = Add-VisioObjectPools $PoolObj $Pool
				Draw_Pool
				$ObjectNumber++
				$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
				$Farm_to_Remote_Desktop_Services_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$PoolNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Pool object " $PoolNumber " of " $PoolTotal " - " $Pool.Name
				}
				Connect-VisioObject $FarmObject $PoolObject
				$y += 1.50
				
				ForEach ( $RDSServer in ( $RDSServersImport | Sort-Object Name | Where-Object { $Farm.Id -eq $_.Farm_Id -and $Farm.RdsServer_Id.contains( $_.Id ) } ) ) `
				{ `
					$x += 2.00
					$RDSHObject = Add-VisioObjectRDSH $RDSHObj $RDSServer
					Draw_RDSH
					$ObjectNumber++
					$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
					$Farm_to_Remote_Desktop_Services_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		
					if ( $debug -eq $true ) `
					{ `
						$RDSServerNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing RDS Server object " $RDSServerNumber " of " $RDSServerTotal " - " $RDSServer.Name
					}
					Connect-VisioObject $PoolObject $RDSHObject
					$PoolObject = $RDSHObject
				}
			}
			$x = 1.50
			$y += 1.50
			
			ForEach ( $Application in ( $ApplicationsImport | Sort-Object Name | Where-Object { $Farm.Id -eq $_.Farm_Id } ) ) `
			{ `
				$x += 2.00
				$ApplicationObject = Add-VisioObjectApplication $ApplicationObj $Application
				Draw_Application
				$ObjectNumber++
				$Farm_to_Remote_Desktop_Services_Complete.Forecolor = "Blue"
				$Farm_to_Remote_Desktop_Services_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add( $Farm_to_Remote_Desktop_Services_Complete )
		
				if ( $debug -eq $true ) `
				{ `
					$ApplicationNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Application object " $ApplicationNumber " of " $ApplicationTotal " - " $Application.Name
				}
				Connect-VisioObject $FarmObject $ApplicationObject
				$FarmObject = $ApplicationObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Quit()
}
#endregion ~~< Farm_to_Remote_Desktop_Services >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Capture_Folder
{ `
	explorer.exe $CaptureCsvFolder
}
#endregion ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Final_Visio
{ `
	$SaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsd"
	$ConvertSaveFile = $VisioFolder + "\" + $ConnServ + " - Horizon vDiagram - " + "$FileDateTime" + ".vsdx"
	$AppVisio = New-Object -ComObject Visio.Application
	$docsObj = $AppVisio.Documents
	$docsObj.Open( $SaveFile ) | Out-Null
	$AppVisio.ActiveDocument.Pages.Item(1).Delete( 1 ) | Out-Null
	$AppVisio.Documents.SaveAs( $SaveFile )
	$AppVisio.Documents.SaveAs( $ConvertSaveFile ) | Out-Null
	Remove-Item $SaveFile
}
#endregion ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Event Handlers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False | Out-Null

Main