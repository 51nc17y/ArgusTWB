# ArgusTWB
This script takes a Tableau workbook file (*twb) as input and returns an excel
(*xlsx) file with up to five sheets: Fields, Parameters, Dashboards, Actions 
and Misc. The returned Excel file can be used as a jumping off point for a data
dictionary. 

# Instructions
Open the script with a text editor such as notepad and replace the string for fileTWB with the path to your workbook. 
Run the script. 
The XLSX file will be created in the same directory as the TWB file. 

#Known Issues
   - Only reports fields used on worksheets
