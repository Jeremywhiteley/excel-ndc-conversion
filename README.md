# excel-ndc-conversion
Excel functions that will convert a 10-digit FDA NDC to a GS1 GTIN-14

# Usage:

1. Enable developer features in Excel
2. On the *Developer* tab click on the *Editor* button to open the editor.
3. Select *Insert/Module*
4. Paste this code into the module window.
5. The *CalculateGTIN* function will then be available for use in your spreadsheet cells.
6. Cell usage example where C3 would be the ID of a cell needing conversion:
    
    =CalculateGTIN(C3)
    
