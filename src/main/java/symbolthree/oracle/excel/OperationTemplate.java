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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/OperationTemplate.java $
 * $Author: Christopher Ho $
 * $Date: 2/17/17 9:58a $
 * $Revision: 18 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

//~--- JDK imports ------------------------------------------------------------

import java.sql.*;

import java.util.*;

public class OperationTemplate extends Operation implements Constants {
    private String objectType = "";

    public OperationTemplate() {
        super();
    }

    @Override
    public void doOperation() throws EXZException {
        Workbook wb            = super.getWorkbook();
        Sheet    dataSheet     = null;
        String   dataSheetName = null;
        
        dataSheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
        
        if (EXZHelper.isEmpty(dataSheetName)) {
        	// use table name as worksheet name
        	dataSheetName = EXZParams.instance().getValue(TABLE_NAME);
        	EXZParams.instance().setValue(DATA_WORKSHEET, dataSheetName);
        }
        
        if (EXZHelper.isEmpty(dataSheetName)) { 
          throw new EXZException("Please specific worksheet name");
        }
        
        EXZHelper.log(LOG_INFO, "worksheetName is " + dataSheetName);        
        
        try {
            int titleRowNo = EXZParams.instance().getInt(COLUMN_TITLE_ROW);

            dataSheet = wb.getSheet(dataSheetName);
            //dataSheet.setSelected(true);

            Cell                  titleCell  = super.getSpecialCells().get(COLUMN_TITLE_FORMAT);
            CellStyle             titleStyle = titleCell.getCellStyle();
            Row                   row        = dataSheet.createRow(titleRowNo - 1);
            Iterator<TableColumn> itr        = super.getColumnMapping().getColumns().iterator();

            while (itr.hasNext()) {
                TableColumn col = (TableColumn) itr.next();

                EXZHelper.log(LOG_DEBUG, "Put Title " + col.getExeclColumnName() + " to column " + col.getExcelColumnNo());

                Cell cell = row.createCell(col.getExcelColumnNo() - 1);

                cell.setCellStyle(titleStyle);

                if (cell instanceof HSSFCell) {
                    cell.setCellValue(new HSSFRichTextString(col.getExeclColumnName()));
                } else if (cell instanceof XSSFCell) {
                    cell.setCellValue(new XSSFRichTextString(col.getExeclColumnName()));
                }
            }

            EXZHelper.log(LOG_INFO, "object Type is " + objectType);
            
            if (objectType.equals("TABLE")) {
                dataSheet.createFreezePane(1, titleRowNo);
            } else {
                dataSheet.createFreezePane(0, titleRowNo);
            }

            // autoSizeColumn
            int maxColNo = super.getColumnMapping().getMaxExcelColumnNo();

            for (int i = 0; i < maxColNo; i++) {
                dataSheet.autoSizeColumn(i);
            }

            super.setSaveFileRequired(true);
            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.TEMPLATE_SHEET_CREATED", dataSheetName));
        
        } catch (Exception e) {
            throw new EXZException(e);
        }
    }

    @Override
    public void defineTableName() throws EXZException {
        EXZHelper.log(LOG_DEBUG, "Operation Template defineTableName");

        // create a worksheet with the worksheet = TABLE_NAME if not present
        Workbook wb        = super.getWorkbook();
        
        String worksheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
        
        if (EXZHelper.isEmpty(EXZParams.instance().getValue(CUSTOM_QUERY)) &&
        	EXZHelper.isEmpty(worksheetName)) {
            
        	worksheetName = EXZParams.instance().getValue(TABLE_NAME);
        }

        if (EXZHelper.isEmpty(worksheetName)) {
          throw new EXZException(EXZI18N.inst().get("ERR.PARAMETER", TABLE_NAME));          
        }
        
        if (wb.getSheet(worksheetName) == null) {
            
        	wb.createSheet(worksheetName);

        	// set new worksheet to be the first sheet
            EXZHelper.log(LOG_DEBUG, "No of worksheet = " +  wb.getNumberOfSheets());

            Hashtable<String, Integer> sheetOrder = new Hashtable<String, Integer>();
            for (int i=0;i<wb.getNumberOfSheets();i++) {
            	
            	String sheetName = wb.getSheetName(i); 

            	if (sheetName.equals(worksheetName)) {
            		sheetOrder.put(sheetName, Integer.valueOf(0));
            	} else {
            		sheetOrder.put(sheetName, Integer.valueOf(i+1));
            	}
            }
            
            Enumeration<String> enu = sheetOrder.keys();
            while (enu.hasMoreElements()) {
            	String sheetName = (String)enu.nextElement();
            	int orderNo = sheetOrder.get(sheetName).intValue();
            	wb.setSheetOrder(sheetName, sheetOrder.get(sheetName).intValue());
            	EXZHelper.log(LOG_DEBUG, "set worksheet " +  sheetName + " to position " + orderNo);
            }
            
        } else {
            throw new EXZException(EXZI18N.inst().get("ERR.TEMPLATE_SHEET_EXIST", worksheetName));
        }
    }

    @Override
    public void createColumnMapping(Map<String, String> customColumnMap) throws EXZException, SQLException {
        EXZHelper.log(LOG_DEBUG, "Operation Template createColumnMapping");

        // for template creation, no need to do column mapping, but need to check Object Type
        String tableName = EXZParams.instance().getValue(TABLE_NAME);

        objectType = super.getObjectType();

        if (!objectType.equals("TABLE") &&
        	!objectType.equals("VIEW") &&
        	!objectType.equals("MATERIALIZED VIEW") &&
        	!objectType.equals(CUSTOM_QUERY)) {
            throw new EXZException(EXZI18N.inst().get("ERR.TABLE_NAME", tableName));
        }

        String sqlStmt = null;
        
        if (objectType.equals(CUSTOM_QUERY)) {
        	sqlStmt = EXZParams.instance().getValue(CUSTOM_QUERY);
        } else {
        	sqlStmt = "SELECT * FROM " + tableName;
        }
        
        Connection        conn            = DBConnection.getInstance().getConnection();
        ResultSet         rs              = conn.createStatement().executeQuery(sqlStmt);
        ResultSetMetaData rsMetaData      = rs.getMetaData();
        int               numberOfColumns = rsMetaData.getColumnCount();

        for (int i = 1; i <= numberOfColumns; i++) {
            TableColumn tabCol = new TableColumn();

            tabCol.setColumnName(rsMetaData.getColumnName(i));
            tabCol.setColumnNullable((rsMetaData.isNullable(i) == 1)
                                     ? true
                                     : false);

            // set the result column to be the first, and the rest will be shift by 1 if objectType=TABLE
            if (objectType.equals("TABLE")) {
                tabCol.setExcelColumnNo(i + 1);
            } else {
                tabCol.setExcelColumnNo(i);
            }

            String columnType = rsMetaData.getColumnTypeName(i);

            tabCol.setColumnType(columnType);

            if (columnType.equals("VARCHAR2")) {
                tabCol.setColumnSize(rsMetaData.getColumnDisplaySize(i));
            }

            super.getColumnMapping().addColumn(tabCol);
        }

        rs.close();

        if (!super.getColumnMapping().checkMapping()) {
            throw new EXZException("Invalid matching");
        }

        // create the result column and set it to the first column
        if (objectType.equals("TABLE")) {
            String      resultColName = EXZParams.instance().getValue(RESULT_COLUMN_NAME);
            TableColumn resultCol     = new TableColumn();

            resultCol.setExeclColumnName(resultColName);
            resultCol.setResultColumn(true);
            resultCol.setExcelColumnNo(1);
            super.getColumnMapping().addColumn(resultCol);
        }

        //super.addROWIDCol();
        
        super.getColumnMapping().showMapping();

        // create a row of column name at row COLUMN_TITLE_ROW (default 1)
        int titleRowNo = 0;

        if (EXZHelper.isEmpty(EXZParams.instance().getValue(COLUMN_TITLE_ROW))) {
            titleRowNo = 1;
        } else {
            titleRowNo = EXZParams.instance().getInt(COLUMN_TITLE_ROW);
        }

        EXZHelper.log(LOG_DEBUG, "Column title will be created in row " + titleRowNo);
    }
}
