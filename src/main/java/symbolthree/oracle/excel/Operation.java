/******************************************************************************
 *
 * ≡ EXZELLENZ ≡
 * Copyright (C) 2009-2016 Christopher Ho 
 * All Rights Reserved, http://www.symbolthree.com
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 *
 * E-mail: christopher.ho@symbolthree.com
 *
 * ================================================
 *
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/Operation.java $
 * $Author: Christopher Ho $
 * $Date: 2/17/17 9:58a $
 * $Revision: 26 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.streaming.*;

import java.io.*;
import java.sql.*;
import java.util.*;

public class Operation implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/Operation.java 26    2/17/17 9:58a Christopher Ho $";
    private ColumnMapping           mapping = new ColumnMapping();
    private Connection              connection;
    private File                    excelFile;
    private boolean                 saveFileRequired;
    private Hashtable<String, Cell> specialCells;
    private Workbook                workbook;

    public Operation() {}

    public void doOperation() throws EXZException {

        // subclass must implement this method
    }

    public void postOperation() throws EXZException {
    	String datasheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
    	
    	EXZHelper.log(LOG_DEBUG, "PostOperation: DATA_WORKSHEET=" + datasheetName);
    	int sheetNo = workbook.getSheetIndex(datasheetName);
        workbook.setSelectedTab(sheetNo);
        workbook.setActiveSheet(sheetNo);
        
        // delete Parameter worksheet if KEEP_PARAMETER_WORKSHEET is N
        String keepParameterWS = EXZParams.instance().getValue(KEEP_PARAMETER_WORKSHEET);
        EXZHelper.log(LOG_DEBUG, "KEEP_PARAMETER_WORKSHEET=" + keepParameterWS);
        if (keepParameterWS != null && keepParameterWS.equals("N")) {
          int sheetIdx = workbook.getSheetIndex(PARAMETER_WORKSHEET_NAME);
          workbook.removeSheetAt(sheetIdx);
        }
    }
    
    public void defineTableName() throws EXZException {
        int    noOfSheet     = workbook.getNumberOfSheets();
        String tableName     = EXZParams.instance().getValue(TABLE_NAME);
        String datasheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
        String customQuery   = EXZParams.instance().getValue(CUSTOM_QUERY);

        /* 4 different cases: */

//    	int sheetNo = workbook.getSheetIndex(datasheetName);
//    	EXZHelper.log(LOG_DEBUG, "defineTableName: " + datasheetName + " Sheet index is " + sheetNo);
    	
    	// tableName exists, datasheetName exists
        if (!EXZHelper.isEmpty(tableName) && !EXZHelper.isEmpty(datasheetName)) {
            // do nothing
        }

        // tableName exists, datasheetName does not exist
        if (!EXZHelper.isEmpty(tableName) && EXZHelper.isEmpty(datasheetName)) {
            if (noOfSheet > 2) {
                throw new EXZException(EXZI18N.inst().get("ERR.PARAMETER", "DATA_WORKSHEET"));
            } else {
            // use tableName as datasheetName
            datasheetName = tableName;
            }
        }

        // tableName does not exist, datasheetName exists
        if (EXZHelper.isEmpty(tableName) && !EXZHelper.isEmpty(datasheetName) && EXZHelper.isEmpty(customQuery)) {
          throw new EXZException(EXZI18N.inst().get("ERR.PARAMETER", "TABLE_NAME"));
        }

        // tableName does not exist, datasheetName does not exist
        if (EXZHelper.isEmpty(tableName) && EXZHelper.isEmpty(datasheetName)) {
                throw new EXZException(EXZI18N.inst().get("ERR.PARAMETER", "DATA_WORKSHEET"));
        }

        if (workbook.getSheet(datasheetName) == null) {
            throw new EXZException(EXZI18N.inst().get("ERR.SHEET_NOTFOUND", datasheetName));
        }

        EXZParams.instance().setValue(TABLE_NAME, tableName);
        EXZParams.instance().setValue(DATA_WORKSHEET, datasheetName);
        EXZHelper.log(LOG_INFO, "TABLE_NAME = " + tableName);
        EXZHelper.log(LOG_INFO, "DATA_WORKSHEET = " + datasheetName);
    }

    public String getObjectType() throws EXZException {

      String objectType = "";	
      
      if (! EXZHelper.isEmpty(EXZParams.instance().getValue(CUSTOM_QUERY))) {
    	objectType = CUSTOM_QUERY;
      } else {
    	// check table or view
        String tableName  = EXZParams.instance().getValue(TABLE_NAME);
        String sqlStmt    = "SELECT OBJECT_TYPE FROM ALL_OBJECTS WHERE OBJECT_NAME='" + 
                             tableName.toUpperCase() + "' AND OWNER IN ('" + 
        		             EXZParams.instance().getValue(USERNAME).toUpperCase() + "', 'PUBLIC')";

        EXZHelper.log(LOG_DEBUG, "Object Checking SQL = " + sqlStmt);        
        
        try {
            Connection conn = DBConnection.getInstance().getConnection();
            if (conn==null) EXZHelper.log(LOG_DEBUG, "Connection is null !");            
            EXZHelper.log(LOG_DEBUG, "Connection closed? " + conn.isClosed());            	
            
            ResultSet  rs   = conn.createStatement().executeQuery(sqlStmt);

            EXZHelper.log(LOG_DEBUG, "Checking User Object Type for " + tableName);

            while (rs.next()) {
                objectType = rs.getString(1);
            }

            rs.close();
            EXZHelper.log(LOG_DEBUG, "User Object Type of " + tableName + " is " + objectType);

            if (objectType.equals("SYNONYM")) {
                sqlStmt = "SELECT a.object_type FROM ALL_OBJECTS a, ALL_SYNONYMS b WHERE " +
                          "a.OBJECT_NAME  = b.TABLE_NAME AND " +
                          "B.SYNONYM_NAME = '" + tableName.toUpperCase() + "' AND " +
                          "b.TABLE_OWNER  IN (a.OWNER, 'PUBLIC') AND " +
                          "a.object_type != 'SYNONYM'";

                EXZHelper.log(LOG_DEBUG, "SQL = " + sqlStmt);
                rs = conn.createStatement().executeQuery(sqlStmt);

                while (rs.next()) {
                    objectType = rs.getString(1);
                }
                
                rs.close();
                
                EXZHelper.log(LOG_DEBUG, "Object Type of " + tableName + " is " + objectType);                
            }
            
            if (!objectType.equals("TABLE") &&
            	!objectType.equals("VIEW") &&
            	!objectType.equals("MATERIALIZED VIEW")) {
                throw new EXZException(EXZI18N.inst().get("ERR.TABLE_NAME", tableName));
            } else {
                mapping.setObjectType(objectType);
            }
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        }
      }
      
      return objectType;
    }

    public void createColumnMapping(Map<String, String> customColumnMap) throws EXZException, SQLException {
        TableColumn col = new TableColumn();

        // add result column if objectType = TABLE
        String objectType = getObjectType();

        EXZHelper.log(LOG_DEBUG, "Operation createColumnMapping for objectType " + objectType);
        
        if (objectType.equals("TABLE") || objectType.equals("VIEW")) {
            String resultCol = EXZParams.instance().getValue(RESULT_COLUMN_NAME);

            col.setExeclColumnName(resultCol);
            col.setResultColumn(true);
            col.setNeeded(false);
            mapping.addColumn(col);
        }

        // insert the custom column mapping to the master ColumnMapping object first
        Iterator<String> itr = customColumnMap.keySet().iterator();

        while (itr.hasNext()) {
            col = new TableColumn();

            String key = (String) itr.next();

            col.setColumnName(key);
            col.setExeclColumnName((String) customColumnMap.get(key));
            col.setNameMatched(false);
            col.setNeeded(true);
            mapping.addColumn(col);
        }

        Connection conn = DBConnection.getInstance().getConnection();

        // get all column names and properties
        String            tableName       = EXZParams.instance().getValue(TABLE_NAME);
        String            sqlStmt         = null;
        
        if (objectType.equals("TABLE") || objectType.equals("VIEW")) {
        	sqlStmt = "SELECT * FROM " + tableName;
        } else if (objectType.equals("CUSTOM_QUERY")) {
        	sqlStmt = EXZParams.instance().getValue(CUSTOM_QUERY);
        }
        
        ResultSet         rs              = conn.createStatement().executeQuery(sqlStmt);
        ResultSetMetaData rsMetaData      = rs.getMetaData();
        int               numberOfColumns = rsMetaData.getColumnCount();

        for (int i = 1; i <= numberOfColumns; i++) {
            TableColumn tabCol = new TableColumn();

            tabCol.setColumnName(rsMetaData.getColumnName(i));
            tabCol.setColumnNullable((rsMetaData.isNullable(i) == 1)
                                     ? true
                                     : false);

            String columnType = rsMetaData.getColumnTypeName(i);

            tabCol.setColumnType(columnType);

            if (columnType.equals("VARCHAR2")) {
                tabCol.setColumnSize(rsMetaData.getColumnDisplaySize(i));
            }

            mapping.addColumn(tabCol);
        }

        rs.close();

        if (!mapping.checkMapping()) {
            throw new EXZException("Invalid matching");
        }

        // mapping.showMapping();
        // check all excelColumn name exists in the data worksheet
        // use TABLE_NAME worksheet if parameter is not defined
        Sheet dataSheet = null;

        if (workbook instanceof HSSFWorkbook) {
            dataSheet = ((HSSFWorkbook) workbook).getSheet(EXZParams.instance().getValue(DATA_WORKSHEET));
        } else if (workbook instanceof XSSFWorkbook) {
            dataSheet = ((XSSFWorkbook) workbook).getSheet(EXZParams.instance().getValue(DATA_WORKSHEET));
        }

        // if title row is not specified it is default to the first row
        int titleRowNo = 0;

        if (EXZHelper.isEmpty(EXZParams.instance().getValue(COLUMN_TITLE_ROW))) {
            titleRowNo = 1;
        } else {
            titleRowNo = EXZParams.instance().getInt(COLUMN_TITLE_ROW);
        }

        EXZParams.instance().setValue(COLUMN_TITLE_ROW, String.valueOf(titleRowNo));

        // scanning original table column
        Iterator<TableColumn> itr2 = mapping.getColumns().iterator();

        while (itr2.hasNext()) {
            TableColumn tabCol     = (TableColumn) itr2.next();
            String      columnName = tabCol.getExeclColumnName();
            Row         row        = dataSheet.getRow(titleRowNo - 1);
            int         lastCellNo = (int) row.getLastCellNum();

//          match column title in Excel and find out its position
            boolean columnMatched = false;

            for (int i = 1; i <= lastCellNo; i++) {
                String titleName = EXZHelper.readString(workbook, dataSheet, titleRowNo, i);

                // EXZHelper.log(LOG_DEBUG, columnName + ":" + titleName);
                if (columnName.equals(titleName)) {
                    tabCol.setExcelColumnNo(i);
                    EXZHelper.log(LOG_DEBUG, columnName + " matched at column " + i);
                    columnMatched = true;

                    break;
                }
            }

            if (! columnMatched &&
            	! tabCol.isColumnNullable() &&
            	! tabCol.isResultColumn() &&
                ! EXZParams.instance().getValue(OPERATION_MODE).equals(OPERATION_DOWNLOAD)) {
                String errMsg = EXZI18N.inst().get("ERR.COLUMN_NOTFOUND", columnName,
                                                   EXZParams.instance().getValue(DATA_WORKSHEET));

                EXZHelper.log(LOG_WARN, errMsg);

                if (EXZParams.instance().getValue(IGONRE_NOT_NULL_COLUMN).equalsIgnoreCase("N")) {
                    throw new EXZException(errMsg);
                } else {
                    tabCol.setNeeded(false);
                    EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.COLUMN_NOTUSED", columnName));
                }
            }
        }

    }

    protected void addROWIDCol() throws EXZException {
        // add ROWID column if not present for TABLE
        String objectType    = getObjectType();
        
        if (objectType.equals("TABLE")) {
          Workbook wb    = getWorkbook();
          Sheet    sheet = null;
  
          if (wb instanceof HSSFWorkbook) {
              sheet = (HSSFSheet) wb.getSheet(EXZParams.instance().getValue(DATA_WORKSHEET));
          } else if (wb instanceof XSSFWorkbook) {
              sheet = (XSSFSheet) wb.getSheet(EXZParams.instance().getValue(DATA_WORKSHEET));
          }
  
          int titleRowNo  = EXZParams.instance().getInt(COLUMN_TITLE_ROW);
          Row row         = sheet.getRow(titleRowNo - 1);
          int colNum      = 1;
          int rowIDColNum = 0;
  
          while (colNum < row.getLastCellNum() + 1) {
              String colName = EXZHelper.readString(wb, sheet, titleRowNo, colNum);
  
              if (colName.equals("ROWID")) {
                  rowIDColNum = colNum;
              } else if (EXZHelper.isEmpty(colName)) {
                  break;
              }
  
              colNum++;
          }
  
          if (rowIDColNum == 0) {
              rowIDColNum = colNum;
  
              Cell      titleCell  = getSpecialCells().get(COLUMN_TITLE_FORMAT);
              CellStyle titleStyle = titleCell.getCellStyle();
              Cell      cell       = row.createCell(rowIDColNum - 1);
  
              cell.setCellStyle(titleStyle);
              cell.setCellValue("ROWID");
          }
  
          TableColumn tabCol = new TableColumn();
  
          tabCol.setColumnName("ROWID");
          tabCol.setExeclColumnName("ROWID");
          tabCol.setExcelColumnNo(rowIDColNum);
          tabCol.setRowIDColumn(true);
          getColumnMapping().addColumn(tabCol);
          
          EXZHelper.log(LOG_INFO, "ROWID on column " + rowIDColNum);
        }
    	
    }
    
    
    public boolean checkDBConnection() {
        boolean useRunAsMode;

        if (EXZParams.instance().getValue(APPS_RUNAS_MODE).equals(USE_RUNAS_MODE)) {
            EXZHelper.log(LOG_DEBUG, "Use RunAs User mode");
            useRunAsMode = true;
        } else {
            useRunAsMode = false;
        }

        if (EXZParams.instance().getValue(CONNECTION_MODE).equals(CONNECTION_DIRECT)) {
            EXZHelper.log(LOG_DEBUG, "Using JDBC URL:" + EXZParams.instance().getJDBUrl());

            try {
                EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.DB_CONNECTING"));
                connection = DBConnection.getInstance(EXZParams.instance().getJDBUrl(),
                             EXZParams.instance().getValue(USERNAME), EXZParams.instance().getValue(PASSWORD),
                            useRunAsMode).getConnection();

                if (EXZParams.instance().getValue(APPS_RUNAS_MODE).equals(USE_RUNAS_MODE)) {}

                EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.DB_CONNECT_SUCCESS"));

                return true;
                
            } catch (EXZException sqle) {
                EXZHelper.logError(sqle);

                return false;
            }
        
        } else if (EXZParams.instance().getValue(CONNECTION_MODE).equals(CONNECTION_EBS)) {
            String dbcFileStr = EXZParams.instance().getValue(DBC_FILE);
            File   dbcFile;

            if (!EXZHelper.isEmpty(dbcFileStr)) {
                dbcFile = new File(dbcFileStr);

                if (!dbcFile.exists() ||!dbcFile.isFile()) {
                    EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("MSG.ERR_INVALID_DBC", dbcFile.getAbsolutePath()));

                    return false;
                }
        
            } else {
                EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("MSG.ERR_INVALID_DBC", dbcFileStr));

                return false;
            }

            try {
                 connection = DBConnection.getInstance(EXZParams.instance().getValue(USERNAME),
                              EXZParams.instance().getValue(PASSWORD), dbcFile, null, useRunAsMode).getConnection();
                EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.DB_CONNECT_SUCCESS"));

                return true;
            } catch (EXZException sqle) {
                EXZHelper.logError(sqle);

                return false;
            }
        } else {
            return false;
        }
    }

    public void writeWorkbook() throws Exception {
        File saveFile;
       
        //if (EXZProp.instance().getBoolean(SAVE_NEW_FILE)) {
        if (EXZParams.instance().getValue(SAVE_NEW_FILE).equals("Y")) {
          saveFile = EXZHelper.getNewFile(excelFile);
        } else {
          saveFile = new File(excelFile.getAbsolutePath());
        }

        // save to a temp file first, then rename / replace the file 
        File tempFile = File.createTempFile("_" + saveFile.getName(), null);
        FileOutputStream fos = new FileOutputStream(tempFile);
        EXZHelper.log(LOG_DEBUG, "Writing file " + tempFile.getAbsolutePath());
        
        if (workbook instanceof HSSFWorkbook) {
            workbook.write(fos);        	
        } else {
        	SXSSFWorkbook sxssfWB = new SXSSFWorkbook((XSSFWorkbook)workbook, SXSSF_WINDOW_SIZE);
        	sxssfWB.write(fos);
        	sxssfWB.close();
        }
        fos.close();

        saveFile.delete();
        tempFile.renameTo(saveFile);
        
        //if (EXZProp.instance().getBoolean(SAVE_NEW_FILE)) {
        if (EXZParams.instance().getValue(SAVE_NEW_FILE).equals("Y")) {
          EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.FILE_CREATED", saveFile.getAbsolutePath()));
        } else {
          EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.FILE_SAVED", saveFile.getAbsolutePath()));
        }
    }
    
    public void setExcelFile(File _file) {
        this.excelFile = _file;
    }

    public File getExcelFile() {
        return excelFile;
    }

/*
    public void setColumnMapping(ColumnMapping _mapping) {
            mapping = _mapping;
    }
*/
    public ColumnMapping getColumnMapping() {
        return mapping;
    }

    public void setSpecialCells(Hashtable<String, Cell> _specialCells) {
        this.specialCells = _specialCells;
    }

    public Hashtable<String, Cell> getSpecialCells() {
        return specialCells;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setConnection(Connection connection) {
        this.connection = connection;
    }

    public Connection getConnection() {
        return connection;
    }

    public void setSaveFileRequired(boolean saveFileRequired) {
        this.saveFileRequired = saveFileRequired;
    }

    public boolean isSaveFileRequired() {
        return saveFileRequired;
    }

    public int toInt(Object obj) {
        if (obj == null) {
            return Integer.MIN_VALUE;
        }

        Integer integer = (Integer) obj;

        return integer.intValue();
    }
    
    public boolean createROWIDFlag() {
    	
    	String val = EXZParams.instance().getValue(CREATE_ROWID_COLUMN);
    	EXZHelper.log(LOG_DEBUG, CREATE_ROWID_COLUMN + " = " + val);
        if (val.equals("Y")) {
        	return true;
        } else {
        	return false;
        }
    }
}
    
