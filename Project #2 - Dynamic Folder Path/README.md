## 📑 Project 2: Getting data from SharePoint Folder via Excel Parameters in Power Query

Externalise the SharePoint folder path in an Excel table as a dynamic parameter, allowing users to update the data sources easily without editing the M code in Power Query Advanced Editor.

## Problem Statement  
When a SharePoint folder name or site structure changes, hardcoded file paths in Power Query break, it requiring manual updates inside the Advanced Editor. This approach is prone to errors when maintaining shared reports.  

## Solutions & Data Structure  
**Step 1:** Create a table in one of the sheets in Excel. The table will a store the target SharePoint folder path.  
- Set up a two-column table in Excel with headers "Parameter" and "File Path".
- Enter TargetFolder under the Parameter column and SharePoint folder path under the File Path column.
- Go to Table Design tab, and rename the table to FPath.

| Parameter | File Path |
| -------- | -------- |
| TargetFolder | /Shared Documents/General/Data Collection Files/ |

<br>

**Step 2:**  Import FPath table into Power Query to create a parameter query.  
- Open the Advanced Editor and replace the generated code with the following M code, then save and exit.
- This query now acts as a dynamic parameter containing the SharePoint path.

```Power Query M
let
    Source = Excel.CurrentWorkbook(){[Name="FPath"]}[Content],
    TargetFolder = Source{0}[Value]
in
    TargetFolder
```

<br>

**Step 3:** Exit Power Query Editor, then Get Data from SharePoint folder as usual.
-	Open the Advanced Editor and replace the generated code of the SharePoint folder path with the following M code.
-	Then save it and continue to locate the required file from the SharePoint folder. Close the Power Query Editor and continue.
-	The folder path can be changed or updated in "FPath" table each time without opening the Power Query Editor.


```Power Query M
let
  TargetFolder = FolderPath,
  Source = SharePoint.Files("https://companyname.sharepoint.com/sites/itteam", [ApiVersion = 15]),
  FilteredFiles = Table.SelectRows(Source, each Text.Contains([Folder Path], TargetFolder))
in
  FilteredFiles
```

