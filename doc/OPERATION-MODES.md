# Operation Modes
- You need to create a template worksheet in order to carry on the other operations (DOWNLOAD, INSERT, UPDATE, or DELETE).  Enter the database info and table name, select the OPERATION_MODE to TEMPLATE.  Run the file in EXZELLENZ, and a new worksheet will be created. The template worksheet name is default to the table / view name.  You could change this name but you must specify the new name in the DATA_WORKSHEET parameter.

- In the template worksheet, you can delete any columns (except RESULT column, which is available for table, but not for view) you don't need.

- The column headings can be changed to more meaningful names, e.g. LAST_NAME to "LastName".  However, you need to put these mappings in the COLUMN_MAPPING parameter so that the program understands which Excel columns are matched to which table columns.

- For UPDATE and DELETE operations, you have to copy the cell belonging to PENDING_UPDATE and PENDING_DELETE to the RESULT column in the template worksheet. Only these "flagged" Excel data rows will be picked up for process.


