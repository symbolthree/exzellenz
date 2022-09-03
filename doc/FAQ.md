<span style="font-size:36px;">FAQ</span><span style="padding-left: 300px;text-align:right;font-size:14px"><a href="INDEX.md">Index</a></span>

---

### What Data Type it can handle ?

VARCHAR2, INTEGER, NUMBER, DATE.  The DATE data is in the format as specified by the DATA_MASK parameter. RAW / CLOB / BLOB are not supported.

### What is CREATE_ROWID_COLUMN parameter used for ?

During TEMPLATE operation, all columns are populated, but ROWID is not in there. In DOWNLOAD and INSERT operations, the ROWID column is then created.  Then you do UPDATE and DELETE operations on this data, this ROWID will be used to identify the data rows. 

If you're going to do INSERT or DOWNLOAD operations and no more subsequent actions, you can simply set this parameter to N.

### What is IGNORE_NOT_NULL_COLUMN means ?

Some columns are defined as "not null", but these columns are not in the template worksheet.  During INSERT operations, if this parameter is "N", an error says "Cannot find column XXX in worksheet YYY".

Some tables uses trigger to populate some non-null column, e.g. primary key value is based on a sequence and populated by a insert trigger.  For this situation, you can turn this parameter to "Y" in order to by-pass this checking.

### What does COLUMN_TITLE_ROW stand for ?

The default value is 1.  If you want to create a customized Excel layout that data is stored in other rows, set this value and during TEMPLATE operation the column titles are created in that row.  In other operations, data is scanned from this row and below.

### Can I remove the RESULT column created in TEMPLATE operations ?

You can remove this column if you only need to do DOWNLOAD operation for this template file.  Also database VIEW does not create RESULT column.  You need this column to store the result message for INSERT, UPDATE and DELETE operations. 

### How the ERROR_HANDLING parameter affects operation ?

COMMIT_AND_EXIT  - If a row of data has error, the process stops at this row and success data is committed to database 
NO_COMMIT_AND_EXIT - If a row of data has error, no data will be committed 
CONTINUE_ON_ERROR - If a row of data has error, this row is skipped and success data is committed to databases 
 
### I want to create an Excel report periodically and distribute it to different parties.  What settings I can use ?

Create a template from a database view. Remove RESULT column.  Set CREATE_ROWID_COLUMN, CONFIRM_OPERATION, KEEP_PARAMETER_WORKSHEET to "N". Set the NEW_FILE_NAME to a name containing a timestamp, e.g. $D{YYYY-MM-DD}.  Run this template using command line (EXZELLENZ.bat).  Then distribute the result file by other script.

### How can I turn on the debug mode to see more details if I encounter an error ?

In GUI mode, from the pop-up menu you can change the Log Level to Debug.  Or you can change the paremeter LOG_LEVEL in the configuration file EXZ.properties. 