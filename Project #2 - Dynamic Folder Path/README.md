## 📑 Project 2: Getting data from SharePoint Folder via Excel Parameters in Power Query

Externalise the SharePoint folder path in an Excel table as a dynamic parameter, allowing users to update the data sources of from the Power Query Advanced Editor without editing the M code.

## Problem Statement  
When the name of folder in SharePoint changed, the query code to extract the file using Power Query will break and it requires manual edit on the statis value of the query code. An Excel table is created to store the file path as parameter, allowing changes easily.  
<br>
When a SharePoint folder name or site structure changes, hardcoded file paths in Power Query break, requiring manual updates inside the Advanced Editor. This approach is prone to errors when maintaining shared reports.  

## Solutions & Data Structure  
**Step 1:** Create a table in one of the sheets in Excel. The table will a store the target SharePoint folder path.  
- Set up a two-column table in Excel with headers Parameter and File Path.
-	Enter TargetFolder under the Parameter column and SharePoint folder path under the File Path column.
-	Go to Table Design tab, and rename the table to FolderPath.

| Parameter | File Path |
| -------- | -------- |
| TargetFolder | /Shared Documents/General/Data Collection Files/ |

<br>

**Step 2:**  Import the FolderPath table into Power Query and create a parameter query.  
-	Open the Advanced Editor and replace the generated code with the following M code, then save and exit.
-	This query now acts as a dynamic parameter containing SharePoint path.

```Power Query M
let
    Source = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content],
    TargetFolder = Source{0}[Value]
in
    TargetFolder
```

<br>

**Step 3:** Connect to SharePoint folder as usual via Get Data from SharePoint Folder.
-	Open the Advanced Editor and replace the generated code of the SharePoint path with the following M code.
-	Then save it and continue to locate the required file from SharePoint Folder.
-	The folder path can be changed or updated in "FolderPath" table each time without opening the Power Query Editor.


```Power Query M
let
  TargetFolder = FolderPath,
  Source = SharePoint.Files("https://companyname.sharepoint.com/sites/itteam", [ApiVersion = 15]),
  FilteredFiles = Table.SelectRows(Source, each Text.Contains([Folder Path], TargetFolder))
in
  FilteredFiles
```

