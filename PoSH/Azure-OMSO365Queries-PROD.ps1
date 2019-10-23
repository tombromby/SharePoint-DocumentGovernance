
<#
 Azure_OMSO365Queries.ps1

  - Synopsis: This script is used to populate datasets for Power BI Reporting on SharePoint Online Activity
 OMS queries are made against the OMS workspace and 

  - Author: Tom Bromby, Kloud Solutions, June 2017
   - Updated 1/11/2017

- requires the following Azure Automation Variables:
	OMSStorageAccount
	OMSStorageAccountKey
	OMSResourceGroup
	OMSWorkspace
	spoAdmin	


 - requires the following Azure Automation modules:
		CSOMforAutomation (custom module with C# assemblies)
		Microsoft.Online.SharePoint.PowerShell (integration module)
		LogAnalyticsQuery (Integration Module)


	Will run on schedule

#>



$cred  = "AzureRunAsConnection"
try
{
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection=Get-AutomationConnection -Name $cred         

    "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 

        "0"
}
catch {
    if (!$servicePrincipalConnection)
    {
        $ErrorMessage = "Connection $cred not found."
        throw $ErrorMessage
    } else{
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}

############################################################################################################################
############################################################################################################################

#Setup variables for Storage and Queries
$StorageAccountName = Get-AutomationVariable -Name 'OMSStorageAccountName'
$StorageAccountKey = Get-AutomationVariable -Name 'OMSStorageAccountKey'
$sactx = New-AzureStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey
$now = Get-Date
$EndDateAndTime = $now.ToString("yyyy-MM-ddTHH:mm:ss")

#Get Workspacename and resource group name
$WorkSpaceName =Get-AutomationVariable -Name 'OMSWorkSpaceName'
$ResourceGroupName = Get-AutomationVariable -Name 'OMSResourceGroupName'
$subscriptionID = $servicePrincipalConnection.SubscriptionId

#Get credentials for RunAs accounts
$spoAdmin = Get-AutomationPSCredential -Name 'spoRunAs' #O365 user account for looking up SPO data




############################################################################################################################
############################################################################################################################



Add-CSOM # Load our Client Side Object Model assemblies for SPO direct API calls - custom module CSOMforAutomation

Import-Module LogAnalyticsQuery #Load the PowerShell module to interact with Azure Log Analytics API

############################################################################################################################
# All the functions for querying SPO

function spoOnlineSiteQuery ($queryName, $siteURI) 
	# This function relies on the Microsoft.Online.SharePoint.PowerShell module loaded into the Automation account
	# This query pulls out all details fo the team sites
	# used for Site Retirement query as it will show the last content change date for the site

	{
			# prepare blob storage variables
			$filename = $queryName + ".csv"
			$StorageContainerName = $query


			# connect to SPO Online
			Connect-SPOService -Url $siteURI -Credential $spoAdmin

			# run query and export to csv
			Get-SPOSite -Detailed -Limit ALL | export-csv $filename -NoTypeInformation 

			# Copy results to blob			
			Set-AzureStorageBlobContent -File $filename -Container $queryName -Blob $filename -Context $sactx -Force

	} #End function



############################################################################################################################


function spoOnlineSiteOwnerQuery ($queryName, $siteURI, $omsQuery) 

# This function relies on the Microsoft.Online.SharePoint.PowerShell module loaded into the Automation account
# This function checks all sites that have had a 'removed from group' operation in the last year, and returns all site owners


	{
			# prepare blob storage variables for output
			$filename = $queryName + ".csv"
			$StorageContainerName = $query


			# connect to SPO Online
			Connect-SPOService -Url $siteURI -Credential $spoAdmin

			
			# Get the copy of the supporting OMS query from Azure storage account
			# Copy CSV with all historical match results locally for updating then we'll upload it back to blob storage
			# blob will be placed in the sandbox, a copy uploaded baclk to blob storage, and the local copy cleanedup once this runbook completes
			$omsBlobName =  $omsQuery + ".csv"
			$omsContainerName = $omsQuery
			$PathToPlaceBlob = "C:\"		
	
			$omsBlob = Get-AzureStorageBlobContent -Blob $omsBlobName -Container $omsContainerName -Destination $PathToPlaceBlob -Context $sactx -Force 
 
				try { 
					Get-Item -Path "$PathToPlaceBlob\$omsBlobName" -ErrorAction Stop 
				} 
				catch { 
					Get-Item -Path $PathToPlaceBlob 
				} 
			
		    $exportCreated =$False
			
			$CSV = import-csv "$PathToPlaceBlob\$omsBlobName" | select UserId, Site_Url, TimeGenerated, Event_Data |`
				ForEach-Object {
					$UserId=$_.UserId
					$siteUrl=$_.Site_Url
					$TimeGenerated = $_.TimeGenerated
					$EventData = $_.Event_Data
					
					$sitegroups = Get-SPOSiteGroup -Site $siteURL

					foreach ($sitegroup in $sitegroups) #for each search result in the OMS query we'll return who is the site owners (and process the result in POwer BI)
									{
										foreach($role in $sitegroup.Roles)
										{
											if ( $role.Contains("Site Owner") -or $role.Contains("Full Control") ) {
													if ($exportCreated) {
																$sitegroup.Title + "`t" + $sitegroup.Users | add-content -path $filename
														}
												if (!$exportCreated)
													{ #if this is the first record to record, create the export file
															$sitegroup.Title + "`t" + $sitegroup.Users | export-csv $filename -NoTypeInformation
															$exportCreated=$True #this is the first record, all the rest we'll append
													}

													}	# end IF												
										} # End FOrEach
									} # End FOrEach
								} # End FOrEach

			# Copy results to blob			
			Set-AzureStorageBlobContent -File $filename -Container $queryName -Blob $filename -Context $sactx -Force

	} #End function


############################################################################################################################


function spoQuery ($spoQuery, $omsQuery) {
		#security classification report - output will be 

		$currentUrl="null" # initialise Site URL parameter
		$CriticalCount = 0
		$DocCount = 0
    	$Skip_URL=$False
		$Critical = $False
		$Classified = $False
		$WriteRecord = $False
		

		#read in OMS query

		# Copy CSV with today's search results locally for uploading and feeding into query
		# blob will be placed in the sandbox and cleanedup once this runbook completes

		$omsBlobName =  $omsQuery + ".csv"
		$omsContainerName = $omsQuery

		# Copy CSV with all historical match results locally for updating then we'll upload it back to blob storage
		# blob will be placed in the sandbox, a copy uploaded baclk to blob storage, and the local copy cleanedup once this runbook completes
		$spoBlobName =  $spoQuery + ".csv"
		$spoContainerName = $spoQuery

		$PathToPlaceBlob = "C:\"		
	
        $omsBlob = Get-AzureStorageBlobContent -Blob $omsBlobName -Container $omsContainerName -Destination $PathToPlaceBlob -Context $sactx -Force 
	    $spoBlob = Get-AzureStorageBlobContent -Blob $spoBlobName -Container $spoContainerName -Destination $PathToPlaceBlob -Context $sactx -Force 
 
				try { 
					Get-Item -Path "$PathToPlaceBlob\$omsBlobName" -ErrorAction Stop 
					Get-Item -Path "$PathToPlaceBlob\$spoBlobName" -ErrorAction Stop 
				} 
				catch { 
					Get-Item -Path $PathToPlaceBlob 
				} 

            $CSV = import-csv "$PathToPlaceBlob\$omsBlobName" | select UserId, OfficeObjectId, Site_Url, TimeGenerated |`
            ForEach-Object {
					$UserId=$_.UserId
					$fileUrl=$_.OfficeObjectId
					$siteUrl=$_.Site_Url
					$Time = $_.TimeGenerated
	
					$filePath=$fileUrl.Replace("https://acme.sharepoint.com","") # trim domain name from start of string
					$DocCount = $DocCount + 1
					$ErrorActionPreference= 'silentlycontinue' # we want to suppress all the 401 errors
							
						
					if ($siteUrl -ne $currentUrl) # we dont want to login to the same site for each record, so only login if weve come to a new site 
							{
									$currentUrl=$siteUrl
									$Skip_URL=$False
									$UserName = $spoAdmin.UserName
									$SecPwd = $spoAdmin.Password
									$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($currentUrl) 
									$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecPwd) 
									$ctx.credentials = $credentials
									$ctx.Load($ctx.Web)
									$ctx.ExecuteQuery()
									
												if ($Error){
														$Skip_URL=$True
														$Error.clear(); #Clear errors so we can check again on the next URL
															}
								}
				
						#if we're not skipping this URL, investigate the metadata of the document
				
						if (!$Skip_URL) {        
								$file = $ctx.Web.GetFileByServerRelativeUrl($filePath);
								$listitem=$file.ListItemAllFields;

								$ctx.Load($file.ListItemAllFields)
								$ctx.ExecuteQuery()

								if ($listitem["stgSecurityClassification_0"] -like "Highly Confidential*") # creates record of each file that is highly sensitive
										{
											$Classified = $True
											$WriteRecord = $True
										}
								if ($listitem["stgCriticalDocument"]) 
										{
											$Critical = $True	
											$WriteRecord = $True
										}
								if ($WriteRecord)	#we've flagged we want to record this record
										{
											"{0},{1},{2},{3},{4},{5}" -f $Time,$UserId,$fileUrl,$siteUrl,$Critical,$Classified | add-content -path "$PathToPlaceBlob\$spoBlobName"
										}
								
							#Clear our variables for the next passthrough
								$Critical = $False
								$Classified = $False
								$WriteRecord = $False
							} #Close IF

				} #Close ForEach
				
                # Copy back the CSV to blob storage with the updated records appended, and we're done - Power BI will do the rest

                Set-AzureStorageBlobContent -File "$PathToPlaceBlob\$spoBlobName" -Container $spoContainerName -Blob $spoBlobName -Context $sactx -Force	

} #End function

##################################################################################### 	




function getAllListItems($_ctx, $_listName, $_rowLimit)

		# This function reads in the list, and for each item on the list (Asset Site), calls a function to list the libraries for that site

	{
		# Load the list
		$lookupList = $_ctx.Web.Lists.GetByTitle($_listName)
		$_ctx.Load($lookupList)

		# Prepare the query
		$query = New-Object Microsoft.SharePoint.Client.CamlQuery
		$query.ViewXml = "<View>
			<RowLimit>$_rowLimit</RowLimit>
		</View>"

		# An array to hold all of the ListItems
		$items = @()

		# Get Items from the List until we reach the end
		do
		{
			$listItems = $lookupList.getItems($query)
			$_ctx.Load($listItems)
			$_ctx.ExecuteQuery()
			$query.ListItemCollectionPosition = $listItems.ListItemCollectionPosition

			foreach($item in $listItems)
			{
				Try
				{
				  Get-SPOWeb $asset
				  $items += $item
				}
				Catch [System.Exception]
				{
					# This shouldn't happen, but just in case
					Write-Host $_.Exception.Message
				}
			}
		}
		While($query.ListItemCollectionPosition -ne $null)

		#return $items
		$_ctx.Dispose()
	} #End function



##################################################################################### 




Function Get-SPOList($Web)
   
		#Function to Get all lists from the web

    {
        #Get All Lists from the web
        $Lists = $Web.Lists
        $ctx.Load($Lists)
        $ctx.ExecuteQuery()
		
		$queryName = 'docgov-sharepoint-librarycreation'
		$fileName = $queryName + '.csv'

        #Get all lists from the web  
        ForEach($List in $Lists)
        {
        #if($List.AllowContentTypes -eq $true)    
        if($List.BaseTemplate -eq "101" -and $list.AllowContentTypes -eq $true)          
            {
                $Web.URL + "," + $List.Title + "," + $List.Created | Out-File -FilePath $fileName -Append
        }
        }
		
		Set-AzureStorageBlobContent -File $fileName -Container $queryName -Blob $fileName -Context $sactx -Force
    } #End function

#####################################################################################

 

    Function Get-SPOWeb($WebURL)

	    #Function to get all sub-webs from given URL

    {
        #Set up the context
        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($WebURL)
        $ctx.Credentials = $Credentials
 
        $Web = $ctx.Web
        $ctx.Load($web)
        #Get all immediate subsites of the site
        $ctx.Load($web.Webs) 
        $ctx.executeQuery()
  
        #Call the function to Get Lists of the web
        #Write-host "Processing Web :"$Web.URL # e.g. Processing Web : https://acmenet.sharepoint.com/sites/IT
        Get-SPOList $Web
  
        #Iterate through each subsite in the current web
        foreach ($Subweb in $web.Webs)
        {
            #Call the function recursively to process all subsites underneaththe current web
            Get-SPOWeb($SubWeb.URL)           
        }
    } #End function
 

##################################################################################### 

		#### QUERIES ####


	#AZURE LOG QUERY
	# SharePoint new files today
	###################################
	# this query data is then used by Check Document Metadata query
	# 28-11-2017 - added fix to filter OfficeObjectId that may contain commas as these break stuff
	$queryName = 'docgov-sharepoint-newfiles-today'
    $fileName =$queryName + '.csv'

	$queryDays = 'P1D'
	$omsQuery = 'OfficeActivity | where OfficeWorkload == "SharePoint" and (Operation=="FileUploaded" or Operation=="FileSyncUploadedFull") and Site_Url startswith "https://acmenet.sharepoint.com/teams" and OfficeObjectId !contains "," | sort by Site_Url asc | project UserId, OfficeObjectId, Site_Url, TimeGenerated'

	$query = Invoke-LogAnalyticsQuery -WorkspaceName $WorkSpaceName -SubscriptionId $subscriptionID -ResourceGroup $ResourceGroupName -Query $omsQuery -Timespan $queryDays -IncludeTabularView

    $query.Results | export-csv $fileName -notypeinformation 
	# Get the query result, export to file in the local automation context
		
	Set-AzureStorageBlobContent -File $fileName -Container $queryName -Blob $fileName -Context $sactx -Force
    # Upload the csv to Azure Storage





	#AZURE LOG QUERY
	# docgov-sharepoint-siteowner-year
	###################################
	# This queries finds all events related to removing site admins or groups, this will then tell the Site Ownership query what team sites to check 
    $queryName = 'docgov-sharepoint-siteowner-year'
    $fileName =$queryName + '.csv'

	$queryDays = 'P365D'
	$omsQuery = 'OfficeActivity | where OfficeWorkload == "SharePoint" and Operation == "RemovedFromGroup" and Event_Data contains "owners" and Site_Url startswith "https://acmenet.sharepoint.com/teams" | project TimeGenerated, UserId, OfficeObjectId, Site_Url, Event_Data'

    $query = Invoke-LogAnalyticsQuery -WorkspaceName $WorkSpaceName -SubscriptionId $subscriptionID -ResourceGroup $ResourceGroupName -Query $omsQuery -Timespan $queryDays -IncludeTabularView

    $query.Results | export-csv $fileName -notypeinformation 
	# Get the query result, export to file in the local automation context
		
	Set-AzureStorageBlobContent -File $fileName -Container $queryName -Blob $fileName -Context $sactx -Force
    # Upload the csv to Azure Storage






	#SPO QUERY
	# Check Document Metadata
	$omsQueryName = 'docgov-sharepoint-newfiles-today'
	$spoQueryName = 'docgov-sharepoint-metadata-year'
	spoQuery $spoQueryName $omsQueryName 






	#SPO QUERY
	# Run the Site Retirement query
	$qName = 'docgov-sharepoint-sitedetails'
	$spoAdminURI = 'https://acmenet-admin.sharepoint.com'
	spoOnlineSiteQuery $qName $spoAdminURI
	
	



<# turning these off until the SPO admin account is given rights to read these team sites, otherwise it's just a stream of 'access denied' errors

	# Run the Site Ownership query
	# this query takes the logged events from docgov-sharepoint-siteowner-year, and from this list of team sites will trawl them all and return the list of current owners
	$qName = 'docgov-sharepoint-siteowners'
	$omsQuery = 'docgov-sharepoint-siteowner-year'
	$spoAdminURI = 'https://acmenet-admin.sharepoint.com'
	spoOnlineSiteOwnerQuery $qName $spoAdminURI $omsQuery


	# Library Creation
	# Once this works (perms) it will read the Asset Lists SP list, and then trawl through all the Asset team sites on that list and list out all libraries
	# those libraries will be written into a big csv file
	# Site owners use a workflow that calls a custom app to create lists and libraries in team sites, becuase of this, no logs are created
	# the report owner will just have to do reports based on library name, location and created date as there is no other useful information
	$siteUrl= 'https://acmenet.sharepoint.com/teams/docXchange' # make var, this is whare the Aset Site list (list of asset sites) is kept
	$UserName = $spoAdmin.UserName
	$SecPwd = $spoAdmin.Password
	$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl) 
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecPwd) 
	$ctx.credentials = $credentials
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()
	$ListTitle = "Asset Sites" # make var
	$mQueryRowLimit = 200
	# we want to now get all list items, and use them in other functions to pull out all libraries
	getAllListItems -_ctx $Ctx -_listName $ListTitle -_rowLimit $mQueryRowLimit


#>

 

