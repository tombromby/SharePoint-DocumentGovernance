/*
# Time Site sharing activity
OfficeActivity | where OfficeWorkload =~ "sharepoint" and (Operation == "SharingSet" or Operation == "SharingInvitationCreated") and Site_Url startswith "https://acmenet.sharepoint.com/teams" | summarize AggregatedValue=count() by Site_Url, TimeGenerated, bin(TimeGenerated, 1d)


*/


let AnalyticsQuery =
let Source = Json.Document(Web.Contents("https://api.loganalytics.io/v1/workspaces/541d0dcb9b79/query", 
[Query=[#"query"="OfficeActivity | where OfficeWorkload =~ ""sharepoint"" and (Operation == ""SharingSet"" or Operation == ""SharingInvitationCreated"") and Site_Url startswith ""https://acmenet.sharepoint.com/teams"" | summarize AggregatedValue=count() by Site_Url, TimeGenerated, bin(TimeGenerated, 1d)",#"x-ms-app"="OmsAnalyticsPBI",#"timespan"="P365D",#"prefer"="ai.response-thinning=true"],Timeout=#duration(0,0,4,0)])),
TypeMap = #table(
{ "AnalyticsTypes", "Type" }, 
{ 
{ "string",   Text.Type },
{ "int",      Int32.Type },
{ "long",     Int64.Type },
{ "real",     Double.Type },
{ "timespan", Duration.Type },
{ "datetime", DateTimeZone.Type },
{ "bool",     Logical.Type },
{ "guid",     Text.Type }
}),
DataTable = Source[tables]{0},
Columns = Table.FromRecords(DataTable[columns]),
ColumnsWithType = Table.Join(Columns, {"type"}, TypeMap , {"AnalyticsTypes"}),
Rows = Table.FromRows(DataTable[rows], Columns[name]), 
Table = Table.TransformColumnTypes(Rows, Table.ToList(ColumnsWithType, (c) => { c{0}, c{3}}))
in
Table
in AnalyticsQuery