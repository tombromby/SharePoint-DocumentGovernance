in here is
powershell for running in Azure automation to perform nightly queries against log analytics and also SPO directly
Power BI M language queries for the data produced
CSOM to use in Azure automation to support C# in PoSH for SPO connectivity

Docgov reports are all the PowerBI M queries - update URIs and analytics workspace GUIDs
Add LogAnalyticsQuery as a module in Azure automation
Add Sharepoint.CSOM as a module in Azure automation, this will hold the DLLs that the Azure Automation server will load to connect to SPO
Under PoSH, there is Azure-OMSO365Queries-PROD.ps1 - this is the PowerShell workflow that you schedule nightly, it runs, loads existing csv files in a storage account, and updates them - there may be some variables and strings you need to update to make it work
