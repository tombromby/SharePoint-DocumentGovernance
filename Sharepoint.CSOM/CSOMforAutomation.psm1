<#
	Used to load CSOM C# DLLs for use in automation script
#>
function Add-CSOM {
		
	#This will open up for usage of CSOM in PowerShell.
    Add-Type -Path "C:\Modules\User\CSOMforAutomation\1.0\Microsoft.SharePoint.Client.dll" 
	Add-Type -Path "C:\Modules\User\CSOMforAutomation\1.0\Microsoft.SharePoint.Client.Runtime.dll" 
	Add-Type -Path "C:\Modules\User\CSOMforAutomation\1.0\Microsoft.SharePoint.Client.UserProfiles.dll" 

	Write-Host "CSOM assemblies loaded"
}

