## 📑 Project 3: Automated Data Unit Conversion Using Power Query

A automated ETL (Extract, Transform, Load) pipeline built in Power Query to clean, standardise, and normalise data (this case: water consumption data from different manufacturing plants) with different units of measurement to a single universal unit.

## Problem Statement  
Data collected across various manufacturing plants contained mixed measurement units (e.g., energy in kWh  vs. MJ , mass in lbs  vs. kg , or water in m3, gal , CCF , and L ).  

## Benefits  
This automated data transformation in Power Query using dynamic master reference table with M code lookups to replace Excel formulas.  
Key benefits:  
- Works across any domain requiring unit conversions, including ESG/Sustainability metrics (water, energy, carbon), supply chain logistics (weights, volumes), and financial analytics.
- New or conversion units can be added directly to the reference table without modifying the core Power Query M code.
- Reduces workbook file size and eliminates errors from using Excel formulas.

## Solutions  
1. Create Conversion Reference Table: A master lookup table containing original units, target universal units, and their exact conversion multipliers.
2. Data Ingestion: Ingests raw datasets and the master lookup table into Power Query.
3. Unit Conversion code: Add the below M code to the raw datasets code, it supports to automatically match the original unit with the single unit from master lookup table, and converted unit data is added into new column.

```Power Query M
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "watersource 2 (liters)", each ConversionRate{[#"Original Unit"=[watersourceunit2]]}?[Convert to Universal Unit]?*[watersource2]),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "watersource 3 (liters)", each ConversionRate{[#"Original Unit"=[watersourceunit3]]}?[Convert to Universal Unit]?*[watersource3]),
    #"Added Custom3" = Table.AddColumn(#"Added Custom2", "watersource 4 (liters)", each ConversionRate{[#"Original Unit"=[watersourceunit4]]}?[Convert to Universal Unit]?*[watersource4]),
    #"Added Custom4" = Table.AddColumn(#"Added Custom3", "watersource 5 (liters)", each ConversionRate{[#"Original Unit"=[watersourceunit5]]}?[Convert to Universal Unit]?*[watersource5])
```
