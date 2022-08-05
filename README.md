# Parse-Data
Parse Data is a code created to modify specific data and output a new spreadsheet with the desired modifications.

Input required - a file called ParseData_Input.xlsx
Output - ParseData_Output.xlsx in C:\_ParsedExcelData_

The input file requires 3 columns, SO License Key(s), Document Number, and Date Created. The code will search through each row and seperate any cells (in the SO Licese Key(s) column) that have multiple keys while keeping it tied to the information in Document Number and Date Created.
