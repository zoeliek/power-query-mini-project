## 📑 Project 1: Replaces values using Power Query to resolve data entry errors in Excel 

## Problem Statement
Raw survey responses (e.g. utility data from factories) collected from the users are saved in Excel workbook but contain multiple data errors.
After the survey closed, a correction table is maintained in the same Excel workbook to store the correct values shared by the users. 
Power Query is used to replace the incorrect values without manually editing.

## Data Structure  
**Table 1 (Raw survey data)**: The original dataset collected from users.  
<img width="992" height="230" alt="image" src="https://github.com/user-attachments/assets/c2040416-a3f1-4c10-a139-66f4e6421f68" />  

**Table 2 (Correction data)**: The correct dataset from the users.  
<img width="730" height="95" alt="image" src="https://github.com/user-attachments/assets/d0b016bf-d5fc-4877-b038-f9fb6e90f369" />  

- Year: Reporting year
- ID: Unique identifier
- Factory Name: Factory (user)
- energyusage.value: Energy consumption value (corrected value, if applicable)
- energyusage.unit: Energy consumption unit (corrected unit, if applicable)
- waterusage.value: Water consumption value (corrected value, if applicable)
- waterusage.unit: Water consumption unit (corrected unit, if applicable)

## Benefits  
- **Preserve the raw data and reduce the risk of accidental changes** — The original dataset remains unchanged to avoid unnecessary manual cell-by-cell edits.
- **Selective correction updates for targeted errors** — Year and Factory Name are used as matching keys to ensure corrections are applied data errors.
- **Dynamic transformation** — The M code identifies require correction without hard-coding every possible correction column.
  
## Methodology
**1. Load the raw data into Power Query**  
- The original survey data is maintained in an Excel table and loaded into Power Query.  
- Define the data type for each column in Power Query Editor then save and exit.  

**2. Create and load the correction table**  
- A second table is created to store the corrected information.  
(This table can be placed on the same sheet as the raw data or on a separate sheet.)  
- The correction table should contain "Year", "Factory Name" and the columns that may require correction.  
- Insert the correct value to the relevant field, and keep "null" (empty) for non correction field.  
- Load the table into Power Query and define the data type for each column in the Editor.  

**3. Apply the corrections dynamically**  
- Once both tables have been loaded into Power Query and the data types have been defined, correction code as below is added into Table 1 Query through Advanced Editor.  
- The code will determine which fields need to be updated in Table 1 based on the correction table from Table 2.  

```Power Query M
Table.FromRecords(
        Table.TransformRows(
            Table1, 
            each let 
                a = Table2{[Year = [Year], Factory Name = [Factory Name]]}? 
            in 
                if a is null then 
                    _
                else 
                    Record.TransformFields(
                        _, 
                        Table.ToList(
                            Table.SelectRows(
                                Record.ToTable(a), 
                                each [Value] <> null and [Value] <> ""
                            ), 
                            each {_{0}, (x) => _{1}}
                        ), 
                        1
                    )
        )
    )

```
**4. Result**  
- Final output of Table 1 data is the corrected version of raw dataset.  
- Close Power Query Editor and load the data into a new sheet.  
- To make the query cleaner, instead of adding correction code into Table 1 Query after data types are defined, reference Table 1 into another Table, then add in the correction code.  
