# AnalyzeInExcel
Analyze in Excel for Power BI Desktop

This is an External Tool for Power BI Desktop to connect Excel to the local model hosted in Power BI Desktop.
The goal is to be just one click away from an Excel report. 
While the initial version must be essential, the following versions could extend features, finding a way to keep the one-click experience consistent.

## v1 Open Excel
A button in External Tools opens Excel through an ODC file.
- The user gets an empty PivotTable connected to the same AS instance hosted by Power BI Desktop. 
- If the user creates a more complex report in Excel and saves the file, the connection string is no longer valid as soon as the Power BI Desktop window is closed.
- The next time the user opens Power BI Desktop, the connection is different, and a file saved in Excel does not have a valid connection.
## v2 Create Excel Report
A button in External Tools that creates a specific report in Excel, using PivotTable and/or PivotChart
- The user gets a configured Excel PivotTable connected to the same AS instance hosted by Power BI Desktop. 
- The report could be based on a template, on an Excel macro, or on a dynamic report created by using Office SDK
- If the user saves the file, the connection string is no longer valid as soon as the Power BI Desktop window is closed.
- The next time the user opens Power BI Desktop, the connection is different, and a file saved in Excel does not have a valid connection.
## v3 Open Excel Report
A button in External Tools that opens an existing report in Excel, or create a new one if it is the first connection to that Power BI report.
- The first time the user uses the button on a PBIX file, the user gets a configured Excel PivotTable connected to the same AS instance hosted by Power BI Desktop.
- The initial report could be based on a template, on an Excel macro, or on a dynamic report created by using Office SDK.
- The connection is lost as soon as the Power BI Desktop window is closed.
- The next time the user opens Power BI Desktop and uses the tool, the tool modifies the Excel file restoring a valid connection and then open the file in Excel.
## v4 Excel Add-In to connect to Power BI Desktop
An Excel add-in written in VSTO that opens Power BI Desktop and connects a PivotTable to Power BI Desktop.
- The user can initiate a connection in Excel, opening a specific Power BI Desktop if it is not already open.
- If the Excel file was saved by using the “Open Excel Report” external tool, then the Add-In has additional information to retrieve the Power BI Desktop file to open.
