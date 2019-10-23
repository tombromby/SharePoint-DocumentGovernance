
#
# Azure_OMSO365Queries.ps1
#
#  - Synopsis: This script is used to populate datasets for Power BI Reporting on SharePoint Online Activity
# OMS queries are made against the OMS workspace and 
#
#  - Author: Tom Bromby, Kloud Solutions, June 2017
# - requires the following Azure Automation Variables:
#	OMSStorageAccount
#	OMSStorageAccountKey
#	OMSResourceGroup
#	OMSWorkspace
#
#	Will run on schedule
#
#



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

#Get credentials for RunAs accounts
$spoAdmin = Get-AutomationPSCredential -Name 'spoRunAs' #O365 user account for looking up SPO data


<#
Cycle through all OMS Queries - for each upload to Azure Blob
		1.	docgov-onedrive-filedeletions-daily
		2.	docgov-onedrive-filedeletions-year
		3.	docgov-onedrive-uploads-daily
		4.	docgov-onedrive-uploads-year
		5.	docgov-onedrive-filesharing-daily
		6.	docgov-onedrive-sharingset-year
		7.	docgov-onedriveSPO-filemodified-daily
		8.	docgov-onedriveSPO-filemodified-year
		9.	docgov-sharepoint-externalsharing-60days
		10.	docgov-sharepoint-filecheckall-60days
		11.	docgov-sharepoint-filecheckouts-daily
		12.	docgov-sharepoint-filecheckouts-year
		13.	docgov-sharepoint-filedeletions-daily
		14.	docgov-sharepoint-filedeletions-year
		15.	docgov-sharepoint-filedownloads-daily
		16.	docgov-sharepoint-filedownloads-year
		17.	docgov-sharepoint-filesharing-daily
		18.	docgov-sharepoint-filesharing-year
		19.	docgov-sharepoint-fileupload-daily
		20.	docgov-sharepoint-fileupload-yearly
		21.	docgov-sharepoint-sitecreations-year
		22.	docgov-sharepoint-teamsharing
		23.	docgov-sharepoint-filecheckedout-60days
		24.	docgov-sharepoint-filecheckin-60days
		25.	docgov-sharepoint-filenaming-year
		26.	docgov-sharepoint-newfiles-daily



#>



############################################################################################################################
############################################################################################################################



Add-CSOM # Load our Client Side Object Model assemblies for SPO direct API calls - custom module CSOMforAutomation


# The following function relies on the Microsoft.Online.SharePoint.PowerShell module loaded into the Automation account

function spoOnlineSiteQuery ($queryName, $siteURI) 
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

	}




# The following function relies on the Microsoft.Online.SharePoint.PowerShell module loaded into the Automation account

function spoOnlineSiteOwnerQuery ($queryName, $siteURI, $omsQuery) 
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

													}

												
										}
									}

					


					}

			# Copy results to blob	
		# hasn't worked, probably a permissions thing : Set-AzureStorageBlobContent : Can not find the specified file 'C:\Temp\q4i3kywd.edf\docgov-sharepoint-siteowners.csv'
			Set-AzureStorageBlobContent -File $filename -Container $queryName -Blob $filename -Context $sactx -Force

	}





function spoQuery ($spoQuery, $omsQuery) {
		#security classification report - output will be 

		$currentUrl="null" # initialise Site URL parameter
		$CriticalCount = 0
		$DocCount = 0
    	$Skip_URL=$False
		$Critical = $False
		$Classified = $False
		$WriteRecord = $False
		

		#read in OMS query - replace with has object from OMS direct

		# Copy CSV with today's search results locally for uloading and feeding into query
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
	
					$filePath=$fileUrl.Replace("https://acmenet.sharepoint.com","") # trim domain name from start of string
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



function runQuery ($query, $Days, $dynamicQuery) {
    
	$filename = $query + ".csv"
	$StorageContainerName = $query
    $StartDateAndTime = $now.AddDays($Days).ToString("yyyy-MM-ddTHH:mm:ss")

    $error.clear()
    $result = @{}
    $StartTime = Get-Date


    # OMS Query
  
    # Get Initial response
    $result = Get-AzureRmOperationalInsightsSearchResults -WorkspaceName $WorkSpaceName `
     -ResourceGroupName $ResourceGroupName -Query $dynamicQuery `
     -Start $StartDateAndTime -End $EndDateAndTime -Top 5000
    $elapsedTime = $(get-date) - $StartTime
    Write-Host "Elapsed: " $elapsedTime "Status: " $result.Metadata.Status

    # Split and extract request Id
    $reqIdParts = $result.Id.Split("/")
    $reqId = $reqIdParts[$reqIdParts.Count -1]

    # Poll if pending
    while($result.Metadata.Status -eq "Pending" -and $error.Count -eq 0) 
			{
				$result = Get-AzureRmOperationalInsightsSearchResults -WorkspaceName $WorkSpaceName -ResourceGroupName $ResourceGroupName -Id $reqId
				$elapsedTime = $(get-date) - $StartTime
				Write-Host "Elapsed: " $elapsedTime "Status: " $result.Metadata.Status
			}

		Write-Host "Returned " $result.Value.Count " documents"
		Write-Host $error

		$result.value | ConvertFrom-Json | export-csv $filename -NoTypeInformation 
		Set-AzureStorageBlobContent -File $filename -Container $query -Blob $filename -Context $sactx -Force

    }




	
##################################################################################### 	
	
	# This function reads in the list, and for each item on the list (Asset Site), calls a function to list the libraries for that site
	function getAllListItems($_ctx, $_listName, $_rowLimit)
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
	}



##################################################################################### 



    #Function to Get all lists from the web
    Function Get-SPOList($Web)
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
    }

#####################################################################################

 
    #Function to get all sub-webs from given URL
    Function Get-SPOWeb($WebURL)
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
    }
 

##################################################################################### 

<#

   #  docgov-onedrive-filedeletions-daily
	$queryName = 'docgov-onedrive-filedeletions-daily'
	$queryDays = '-60'
	$omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive Operation=FileDeleted | measure count() by UserId Interval 1DAY"
    runQuery $queryName $queryDays $omsQuery


    # docgov-onedrive-filedeletions-year
	$queryName = 'docgov-onedrive-filedeletions-year'
	$queryDays = '-365'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive (Operation=FileUploaded OR Operation=FileSyncUploadedFull) | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery

   
    # docgov-onedrive-uploads-daily
    $queryName = 'docgov-onedrive-uploads-daily'
	$queryDays = '-60'
	$omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive (Operation=FileUploaded OR Operation=FileSyncUploadedFull) | measure count() by UserId Interval 1DAY"
    runQuery $queryName $queryDays $omsQuery


    # docgov-onedrive-uploads-year
	$queryName = 'docgov-onedrive-uploads-year'
	$queryDays = '-365'	
	$omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive (Operation=FileUploaded OR Operation=FileSyncUploadedFull) | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery

	# docgov-onedrive-filesharing-daily
	$queryName = 'docgov-onedrive-filesharing-daily'
	$queryDays = '-60'	
	$omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive Operation=SharingSet | measure count() by UserId Interval 24HOURS"
    runQuery $queryName $queryDays $omsQuery

	# docgov-onedrive-sharingset-year
	$queryName = 'docgov-onedrive-sharingset-year'
	$queryDays = '-365'	
	$omsQuery = "Type=OfficeActivity OfficeWorkload=onedrive Operation=SharingSet | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery

    # docgov-onedriveSPO-filemodified-daily
	$queryName = 'docgov-onedrivespo-filemodified-daily'
	$queryDays = '-60'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint OR OfficeWorkload=onedrive (Operation=FileModified) | measure count() by OfficeWorkload Interval 1DAY"
    runQuery $queryName $queryDays $omsQuery


    # docgov-onedriveSPO-filemodified-year    
	$queryName = 'docgov-onedrivespo-filemodified-year'
	$queryDays = '-365'
 	$omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint OR OfficeWorkload=onedrive (Operation=FileModified) | measure count() by OfficeWorkload Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery

    # docgov-sharepoint-externalsharing-60days
	$queryName = 'docgov-sharepoint-externalsharing-60days'
	$queryDays = '-60' 
	$omsQuery = "OfficeWorkload=sharepoint Operation=SharingInvitationAccepted | select UserId, Site_Url, TimeGenerated, Event_Data"
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filecheckall-60days
	$queryName = 'docgov-sharepoint-filecheckall-60days'
	$queryDays = '-60'
    $omsQuery = 'Type=OfficeActivity Operation=FileCheck* (SourceFileExtension=xsl* OR SourceFileExtension=doc*) UserId!="spoadmin@acme.com.au"'
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filecheckouts-daily
	$queryName = 'docgov-sharepoint-filecheckouts-daily'
	$queryDays = '-60'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileCheckedOut | measure count() by UserId Interval 24HOURS"
    runQuery $queryName $queryDays $omsQuery

    
    # docgov-sharepoint-filecheckouts-year
	$queryName = 'docgov-sharepoint-filecheckouts-year'
	$queryDays = '-365'
	$omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileCheckedOut | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filedeletions-daily
	$queryName = 'docgov-sharepoint-filedeletions-daily'
	$queryDays = '-60'
	$omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileDeleted SourceFileName!=RegEx("~$@") Site_Url=RegEx("https://acmenet.sharepoint.com/teams@") | measure count() by UserId Interval 24HOURS'
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filedeletions-year
	$queryName = 'docgov-sharepoint-filedeletions-year'
	$queryDays = '-365'
	$omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileDeleted SourceFileName!=RegEx("~$@") Site_Url=RegEx("https://acmenet.sharepoint.com/teams@") | measure count() by UserId Interval 7DAYS'
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filedownloads-daily
	$queryName = 'docgov-sharepoint-filedownloads-daily'
	$queryDays = '-60'
	$omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileDownloaded | measure count() by UserId Interval 24HOURS"
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filedownloads-year
	$queryName = 'docgov-sharepoint-filedownloads-year'
	$queryDays = '-365'
	$omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint Operation=FileDownloaded | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery

    
    # docgov-sharepoint-filesharing-daily
	$queryName = 'docgov-sharepoint-filesharing-daily'
	$queryDays = '-60'
    $omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint Operation=SharingSet Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | measure count() by UserId Interval 24HOURS'
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-filesharing-year
	$queryName = 'docgov-sharepoint-filesharing-year'
	$queryDays = '-365'
	$omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint Operation=SharingSet Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | measure count() by UserId Interval 7DAYS'
    runQuery $queryName $queryDays $omsQuery

    # docgov-sharepoint-fileupload-daily
	$queryName = 'docgov-sharepoint-fileupload-daily'
	$queryDays = '-60'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint (Operation=FileUploaded OR Operation=FileSyncUploadedFull) | measure count() by UserId Interval 1DAY"
    runQuery $queryName $queryDays $omsQuery

    # docgov-sharepoint-fileupload-yearly
	$queryName = 'docgov-sharepoint-fileupload-yearly'
	$queryDays = '-365'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint (Operation=FileUploaded OR Operation=FileSyncUploadedFull) | measure count() by UserId Interval 7DAYS"
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-sitecreations-year 
	$queryName = 'docgov-sharepoint-sitecreations-year'
	$queryDays = '-365'
    $omsQuery = "Type=OfficeActivity OfficeWorkload=sharepoint Operation=SiteCollectionCreated | measure count() by OfficeObjectID_CF Interval 24HOURS"
    runQuery $queryName $queryDays $omsQuery


    # docgov-sharepoint-teamsharing 
	$queryName = 'docgov-sharepoint-teamsharing'
	$queryDays = '-60'
    #$omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint (Operation=SharingSet or Operation=SharingInvitationCreated) Site_Url!="https://acmenet.sharepoint.com/sites/OnlineForms"  UserId!="spoadmin@acme.com.au" | measure count() by Site_Url Interval 1DAY'
    $omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint (Operation=SharingSet or Operation=SharingInvitationCreated) Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | measure count() by Site_Url Interval 1DAY'
	runQuery $queryName $queryDays $omsQuery  
	

    
    # docgov-sharepoint-filecheckedout-60days
	$queryName = 'docgov-sharepoint-filecheckedout-60days'
	$queryDays = '-60'
    $omsQuery = 'Type=OfficeActivity Operation=FileCheckedOut (SourceFileExtension=xsl* OR SourceFileExtension=doc*) UserId!="spoadmin@acme.com.au" | select TimeGenerated,Operation,OfficeObjectId,UserId'
    runQuery $queryName $queryDays $omsQuery 

    
    # docgov-sharepoint-filecheckin-60days
	$queryName = 'docgov-sharepoint-filecheckin-60days'
	$queryDays = '-60'
    $omsQuery = 'Type=OfficeActivity Operation=FileCheckedIn OR Operation=FileCheckOutDiscarded (SourceFileExtension=xsl* OR SourceFileExtension=doc*) UserId!="spoadmin@acme.com.au" Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | select TimeGenerated,Operation,OfficeObjectId,UserId'
    runQuery $queryName $queryDays $omsQuery 


    # docgov-sharepoint-filenaming-year
	$queryName = 'docgov-sharepoint-filenaming-year'
	$queryDays = '-365'
    $omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint (Operation=FileUploaded OR Operation=FileSyncUploadedFull) Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") (SourceFileName=RegEx("@20@&@FINAL@") OR SourceFileName=RegEx("@20@&@DRAFT@")) | measure count() by UserId Interval 1DAY'#Select TimeGenerated, UserId, SourceFileName, Site_Url
    runQuery $queryName $queryDays $omsQuery 
	# Site_Url=RegEx("@/teams/@") 
   
    # docgov-sharepoint-newfiles-today
	$queryName = 'docgov-sharepoint-newfiles-today'
	$queryDays = '-1'
    $omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint (Operation=FileUploaded OR Operation=FileSyncUploadedFull) Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | Select UserId, OfficeObjectId, Site_Url, TimeGenerated | sort Site_Url'
    runQuery $queryName $queryDays $omsQuery
	
	# Check Document Metadata
	$omsQueryName = 'docgov-sharepoint-newfiles-today'
	$spoQueryName = 'docgov-sharepoint-metadata-year'
	spoQuery $spoQueryName $omsQueryName 

	# Run the Site Retirement query
	$qName = 'docgov-sharepoint-sitedetails'
	$spoAdminURI = 'https://acmenet-admin.sharepoint.com'
	spoOnlineSiteQuery $qName $spoAdminURI
	
	
	# docgov-sharepoint-siteowner-year
	$queryName = 'docgov-sharepoint-siteowner-year'
	$queryDays = '-365'
    $omsQuery = 'Type=OfficeActivity OfficeWorkload=sharepoint Operation=RemovedFromGroup Event_Data=owners Site_Url=RegEx("https://acmenet.sharepoint.com/teams/@") | Select TimeGenertated, UserId, OfficeObjectId, Site_Url, Event_Data'
    runQuery $queryName $queryDays $omsQuery
	#>

	# Run the Site Ownership query
	$qName = 'docgov-sharepoint-siteowners'
	$omsQuery = 'docgov-sharepoint-siteowner-year'
	$spoAdminURI = 'https://acmenet-admin.sharepoint.com'
	spoOnlineSiteOwnerQuery $qName $spoAdminURI $omsQuery



	$siteUrl= 'https://acmenet.sharepoint.com/teams/docXchange' # make var
	$UserName = $spoAdmin.UserName
	$SecPwd = $spoAdmin.Password
	$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl) 
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,$SecPwd) 
	$ctx.credentials = $credentials
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$ListTitle = "Asset Sites" # make var
	$mQueryRowLimit = 200

	# we want to now get all list items
	getAllListItems -_ctx $Ctx -_listName $ListTitle -_rowLimit $mQueryRowLimit




	
	write-host "Queries Complete... "    

