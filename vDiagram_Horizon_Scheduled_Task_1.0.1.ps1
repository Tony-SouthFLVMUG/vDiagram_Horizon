<# 
.SYNOPSIS 
   vDiagram Scheduled Export

.DESCRIPTION
   vDiagram Scheduled Export

.NOTES 
   File Name	: vDiagram_Horizon_Scheduled_Task_1.0.1.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 1.01

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data

.CHANGE LOG
	- 09/12/2021 - v1.0.1
		Initial release
#>

#region Variables
$ConnServ = "Replace with Horizon Connection Server name."
$CaptureCsvFolder = "C:\vDiagram\Capture"
$SMTPSRV = "SMTP Server"
$EmailFrom = "outbound@email.com"
$EmailTo = "you@email.com"
$Subject = "vDiagram Horizon 1.0 Files"

# !!!!!!!!!!!! Comment out Line 744 once .xml has been created !!!!!!!!!!!!

# Variables (no need to edit)
$Date = (Get-Date -format "yyyy-MM-dd")
$ScriptPath = (Get-Item (Get-Location)).FullName
$XMLFile = $ScriptPath + "\credentials.xml"
$ZipFile = "$ScriptPath" + "\vDiagram Files" + " " + "$Date.zip"
$AttachmentFile = $ZipFile
$EmailSubject = "vDiagram Horizon 1.0 Files"
$Body = $EmailSubject
#endregion

#region Functions

#region PsCreds

#region Export-PSCredential
Function Export-PSCredential {
        param ( $Credential = (Get-Credential), $Path = "credentials.xml" )
 
        # Look at the object type of the $Credential parameter to determine how to handle it
        switch ( $Credential.GetType().Name ) {
                # It is a credential, so continue
                PSCredential            { continue }
                # It is a string, so use that as the username and prompt for the password
                String                          { $Credential = Get-Credential -credential $Credential }
                # In all other caess, throw an error and exit
                default                         { Throw "You must specify a credential object to export to disk." }
        }
       
        # Create temporary object to be serialized to disk
        $export = "" | Select-Object Username, EncryptedPassword
       
        # Give object a type name which can be identified later
        #$export.PSObject.TypeNames.Insert(0,’ExportedPSCredential’)
       
        $export.Username = $Credential.Username
 
        # Encrypt SecureString password using Data Protection API
        # Only the current user account can decrypt this cipher
        $export.EncryptedPassword = $Credential.Password | ConvertFrom-SecureString
 
        # Export using the Export-Clixml cmdlet
        $export | Export-Clixml $Path
        Write-Host -foregroundcolor Green "Credentials saved to: " -noNewLine
 
        # Return FileInfo object referring to saved credentials
        Get-Item $Path
}
#endregion Export-PSCredential

#region Import-PSCredential 
Function Import-PSCredential {
        param ( $Path = "credentials.xml" )
 
        # Import credential file
        $import = Import-Clixml $Path
       
        # Test for valid import
        if ( !$import.UserName -or !$import.EncryptedPassword ) {
                Throw "Input is not a valid ExportedPSCredential object, exiting."
        }
        $Username = $import.Username
       
        # Decrypt the password and store as a SecureString object for safekeeping
        $SecurePass = $import.EncryptedPassword | ConvertTo-SecureString
       
        # Build the new credential object
        $Credential = New-Object System.Management.Automation.PSCredential $Username, $SecurePass
        Write-Output $Credential
}
#endregion Import-PSCredential 

#endregion PsCreds

#region vCenterFunctions

#region ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Connect_HVServer
{ `
	$global:HorizonViewServer = Connect-HVServer -server $ConnServ -Credential (Import-PSCredential -path $XMLFile)
	$global:HorizonViewAPI = $HorizonViewServer.ExtensionData
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )
}
#endregion

#region ~~< Disconnect_HVServer >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Disconnect_HVServer
{ `
	Disconnect-HVServer -Confirm:$false
}
#endregion

#endregion

#region CsvExportFunctions

#region ~~< VirtualCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VirtualCenter_Export
{ `
	$VirtualCenter_CSV = "$CaptureCsvFolder\$ConnServ-VirtualCenterExport.csv"
	$VirtualCenter = $HorizonViewAPI.VirtualCenter.VirtualCenter_List()
	$VirtualCenterHealth = $HorizonViewAPI.VirtualCenterHealth.VirtualCenterHealth_List()
	
	ForEach ( $VC in $VirtualCenterHealth ) `
	{  `
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
	$ComposerServers_CSV = "$CaptureCsvFolder\$ConnServ-ComposerServersExport.csv"
	$ComposerServers = $HorizonViewAPI.ViewComposerHealth.ViewComposerHealth_List()
	ForEach ( $ComposerServer in $ComposerServers) `
	{ `
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
	}
}
#endregion ~~< ComposerServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectionServers_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ConnectionServers_Export
{ `
	$ConnectionServers_CSV = "$CaptureCsvFolder\$ConnServ-ConnectionServersExport.csv"
	$ConnectionServerHealth = $HorizonViewAPI.ConnectionServerHealth.ConnectionServerHealth_List()
	$ConnectionServers = $HorizonViewAPI.ConnectionServer.ConnectionServer_List()
	ForEach ( $ConnectionServer in $ConnectionServers ) `
	{ `
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
	
	ForEach	( $Pool in $DesktopSummaryView ) `
	{ `
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
	
	ForEach ( $Desktop in $MachineNamesView ) `
	{ `
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
	
	ForEach ( $RDSServer in $RDSServerInfo ) `
	{ `
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

    ForEach ( $Farm in $FarmHealthInfo ) `
	{ `
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
	$Applications_CSV = "$CaptureCsvFolder\$ConnServ-ApplicationsExport.csv"
	$Query_Service = New-Object "Vmware.Hv.QueryServiceService"
	$Query = New-Object "Vmware.Hv.QueryDefinition"

	$Query.QueryEntityType = 'ApplicationInfo'
	$ApplicationInfoQuery = $Query_Service.QueryService_Query( $HorizonViewAPI,$Query )
	$ApplicationInfo = $ApplicationInfoQuery.Results
	$Query_Service.QueryService_DeleteAll( $HorizonViewAPI )

	ForEach ( $Application in $ApplicationInfo ) `
	{ `
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
	$Gateways_CSV = "$CaptureCsvFolder\$ConnServ-GatewaysExport.csv"
	$Gateways = $HorizonViewAPI.GatewayHealth.GatewayHealth_List()
	
	ForEach ( $Gateway in $Gateways ) `
	{ `
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

#endregion

#endregion

#region Export-PSCredential
Export-PSCredential
#endregion

#region Tasks
Connect_HVServer; VirtualCenter_Export; ComposerServers_Export; ConnectionServers_Export; Pools_Export; Desktops_Export; RDSServers_Export; Farms_Export; Applications_Export; Gateways_Export; Disconnect_HVServer
#endregion

#region Zip Files
Compress-Archive -U -Path $CaptureCsvFolder -DestinationPath $ZipFile
#endregion

#region Send E-mail
$msg = new-object Net.Mail.MailMessage
$att = new-object Net.Mail.Attachment($AttachmentFile)
$smtp = new-object Net.Mail.SmtpClient($SMTPSRV) 
$msg.From = $EmailFrom
$msg.To.Add($EmailTo)
$msg.Subject = $EmailSubject
$msg.Body = $Body
$msg.Attachments.Add($AttachmentFile) 
$smtp.Send($msg)
#endregion