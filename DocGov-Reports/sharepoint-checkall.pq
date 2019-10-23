/*
sharepoint-checkin-out
OfficeActivity | where Operation startswith_cs "FileCheck" and (SourceFileExtension startswith_cs "xsl" or SourceFileExtension startswith_cs "doc") and UserId != "spoadmin@acme.com.au" and Site_Url startswith "https://acmenet.sharepoint.com/teams"


*/


let AnalyticsQuery =
let Source = Json.Document(Web.Contents("https://api.loganalytics.io/v1/workspaces/396-541d0dcb9b79/query", 
[Query=[#"query"="OfficeActivity | where Operation startswith_cs ""FileCheck"" and (SourceFileExtension startswith_cs ""xsl"" or SourceFileExtension startswith_cs ""doc"") and UserId != ""spoadmin@stockland.com.au"" and Site_Url startswith ""https://stocklandnet.sharepoint.com/teams""",#"x-ms-app"="OmsAnalyticsPBI",#"timespan"="P365D",#"prefer"="ai.response-thinning=true"],Timeout=#duration(0,0,4,0)])),
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