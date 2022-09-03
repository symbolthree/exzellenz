/******************************************************************************
 *
 * ≡ EXZELLENZ ≡
 * Copyright (C) 2009-2022 Christopher Ho 
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
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.sql.*;
import java.util.*;

public class OperationDownload extends Operation implements Constants {

    private int SHOWING_ROWCOUNT = EXZProp.instance().getInt(EXZ_LOG_INTERVAL);
    
    public OperationDownload() {
        super();
    }

    @Override
    public void doOperation() throws EXZException {
        Workbook                wb            = super.getWorkbook();
        String                  worksheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
        Sheet                   dataSheet     = wb.getSheet(worksheetName);
        String                  tableName     = EXZParams.instance().getValue(TABLE_NAME);
        Hashtable<String, Cell> specialCells  = super.getSpecialCells();
        String                  objectType    = super.getObjectType();
        String                  customSQL     = EXZParams.instance().getValue(CUSTOM_QUERY);

        // create SQL Statement
        
        String selectSQL = null;
        ArrayList<Integer>    excelColPosition = new ArrayList<Integer>();
        ArrayList<String>     colType          = new ArrayList<String>();
        Iterator<TableColumn> itr              = super.getColumnMapping().getColumns().iterator();	        

        selectSQL = "SELECT ";

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            if (!tabCol.isResultColumn()) {
                String columnName = tabCol.getColumnName();

                if (objectType.equals("TABLE")) {
                    colType.add(tabCol.getColumnType());
                    selectSQL = selectSQL + columnName + "," + System.getProperty("line.separator");
                    excelColPosition.add(new Integer(tabCol.getExcelColumnNo()));                    
                }

                if (!objectType.equals("TABLE") && ! columnName.equals("ROWID")) {
                    colType.add(tabCol.getColumnType());
                    selectSQL = selectSQL + columnName + "," + System.getProperty("line.separator");
                    excelColPosition.add(new Integer(tabCol.getExcelColumnNo()));                    
                }
            }
        }

	    if (!objectType.equals(CUSTOM_QUERY)) {
	        	
	        // take out the last ","
	        selectSQL = selectSQL.substring(0, selectSQL.length() - System.getProperty("line.separator").length() - 1);
	
	        // selectSQL = selectSQL + " ROWID ";
	        selectSQL = selectSQL + " FROM " + tableName;
	
	        String whereClause = EXZParams.instance().getValue(WHERE_CLAUSE);
	        if (whereClause != null && whereClause.trim().toUpperCase().startsWith("WHERE ")) {
	        	whereClause = whereClause.substring(5);
	        }
	        String orderClause = EXZParams.instance().getValue(ORDER_CLAUSE);
	        if (orderClause != null && orderClause.trim().toUpperCase().startsWith("ORDER BY ")) {
	        	orderClause = orderClause.substring(8);
	        }
	
	        if (!EXZHelper.isEmpty(whereClause)) {
	            selectSQL = selectSQL + " WHERE " + whereClause;
	        }
	
	        if (!EXZHelper.isEmpty(orderClause)) {
	            selectSQL = selectSQL + " ORDER BY " + orderClause;
	        }

        } else {
        	
        	selectSQL = customSQL;
        }
        
        EXZHelper.log(LOG_DEBUG, "Download Query: " + selectSQL);

        // start writing the cell values
        try {
            int               noOfRow    = 0;
            int               excelRowNo = EXZParams.instance().getInt(COLUMN_TITLE_ROW) + 1;
            Connection        conn       = DBConnection.getInstance().getConnection();
            ResultSet         rs         = conn.createStatement().executeQuery(selectSQL);
            ResultSetMetaData metaData   = rs.getMetaData();
            Row               row;

            EXZHelper.log(LOG_DEBUG, "excelColPosition.size=" + excelColPosition.size());

            EXZHelper.log(LOG_INFO, "Start downloading...");
            
            while (rs.next()) {
                row = dataSheet.createRow(excelRowNo - 1);

                for (int i = 1; i <= excelColPosition.size(); i++) {
                    int  colPos = toInt(excelColPosition.get(i - 1));
                    Cell cell   = null;

                    if (wb instanceof HSSFWorkbook) {
                        cell = (HSSFCell) cell;
                    } else if (wb instanceof XSSFWorkbook) {
                        cell = (XSSFCell) cell;
                    }

                    if (colPos > 0) {
                        cell = row.createCell(colPos - 1);
                        
                        /*** VARCHAR2 ***/
                        if (metaData.getColumnTypeName(i).equals("VARCHAR2")) {
                            String value = rs.getString(i);

                            //cell.setCellValue(Cell.CELL_TYPE_STRING);
                            cell.setCellType(CellType.STRING);
                            
                            if (cell instanceof HSSFCell) {
                                cell.setCellValue(new HSSFRichTextString(value));
                            } else if (cell instanceof XSSFCell) {
                                cell.setCellValue(new XSSFRichTextString(value));
                            }
                        
                        /*** INTEGER ***/
                        } else if (metaData.getColumnTypeName(i).equals("NUMBER") && 
                        		   metaData.getScale(i) == 0 &&
                        		   metaData.getPrecision(i) != 0) {
                            int value = rs.getInt(i);

                            if (!rs.wasNull()) {
                              //cell.setCellValue(Cell.CELL_TYPE_NUMERIC);
                            	cell.setCellType(CellType.NUMERIC);
                              cell.setCellValue(value);
                            }
                            
                        /*** NUMBER ***/    
                        } else if (metaData.getColumnTypeName(i).equals("NUMBER") && 
                        		metaData.getScale(i) < 0) {
                            double value = rs.getDouble(i);

                            if (!rs.wasNull()) {
                                //cell.setCellValue(Cell.CELL_TYPE_NUMERIC);
                            	cell.setCellType(CellType.NUMERIC);
                                cell.setCellValue(value);
                            }

                        /*** NUMBER ***/    
                        } else if (metaData.getColumnTypeName(i).equals("NUMBER") && 
                        		metaData.getScale(i) == 0 &&
                        		metaData.getPrecision(i) == 0) {
                            double value = rs.getDouble(i);

                            if (!rs.wasNull()) {
                                //cell.setCellValue(Cell.CELL_TYPE_NUMERIC);
                                cell.setCellType(CellType.NUMERIC);
                                cell.setCellValue(value);
                            }
                                
                        /*** DATE ***/    
                        } else if (metaData.getColumnTypeName(i).equals("DATE")) {
                            Timestamp value = rs.getTimestamp(i);

                            if (value != null) {
                                CellStyle cellStyle = null;

                                if (cell instanceof HSSFCell) {
                                    cellStyle =
                                        (HSSFCellStyle) ((HSSFCell) specialCells.get("DATE_FORMAT")).getCellStyle();
                                } else if (cell instanceof XSSFCell) {
                                    cellStyle =
                                        (XSSFCellStyle) ((XSSFCell) specialCells.get("DATE_FORMAT")).getCellStyle();
                                }

                                cell.setCellStyle(cellStyle);
                                cell.setCellValue((java.util.Date) value);
                            }
                        } else {
                            String value = rs.getString(i);

                            //cell.setCellValue(Cell.CELL_TYPE_STRING);
                            cell.setCellType(CellType.STRING);

                            if (cell instanceof HSSFCell) {
                                cell.setCellValue(new HSSFRichTextString(value));
                            } else if (cell instanceof XSSFCell) {
                                cell.setCellValue(new XSSFRichTextString(value));
                            }
                        }
                    }
                }

                excelRowNo++;
                noOfRow++;
                
                if (noOfRow%SHOWING_ROWCOUNT == 0) {
                  EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.DOWNLOAD_ROW_DATA", String.valueOf(noOfRow)));
                }
                
                EXZHelper.log(LOG_DEBUG, "No of row created: " + noOfRow);
            }

            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.DOWNLOAD_ROW_DATA", String.valueOf(noOfRow)));

            // autoSizeColumn
            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.FORMAT_SHEET"));            
            int maxColNo = super.getColumnMapping().getMaxExcelColumnNo();

            for (int i = 0; i < maxColNo; i++) {
                dataSheet.autoSizeColumn(i);
            }

            // hide ROWID column
            if (createROWIDFlag()) {            
	            if (objectType.equals("TABLE")) {
	                int rowidColNo = super.getColumnMapping().getROWIDColumn();
	
	                if (!dataSheet.isColumnHidden(rowidColNo - 1)) {
	                    dataSheet.setColumnHidden(rowidColNo - 1, true);
	                }
	            }
            }

            // set focus on this worksheet
            // wb.setSheetOrder(worksheetName, 0);
            // wb.setFirstVisibleTab(wb.getSheetIndex(worksheetName));
            
            super.setSaveFileRequired(true);
            
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        } catch (Exception e) {
            super.setSaveFileRequired(true);
            e.printStackTrace();

            return;
        }
    }

    @Override
    public void defineTableName() throws EXZException {
        super.defineTableName();
    }

    @Override
    public void createColumnMapping(Map<String, String> customColumnMap) throws EXZException, SQLException {
        super.createColumnMapping(customColumnMap);
        if (createROWIDFlag()) {
          super.addROWIDCol();
        }

/*
        // add ROWID column if not present for TABLE
        String objectType    = super.getObjectType();
        
        if (objectType.equals("TABLE")) {
          Workbook wb    = super.getWorkbook();
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
  
              Cell      titleCell  = super.getSpecialCells().get(COLUMN_TITLE_FORMAT);
              CellStyle titleStyle = titleCell.getCellStyle();
              Cell      cell       = row.createCell(rowIDColNum - 1);
  
              cell.setCellStyle(titleStyle);
              cell.setCellValue("ROWID");
          }
  
          TableColumn tabCol = new TableColumn();
  
          tabCol.setColumnName("ROWID");
          tabCol.setExeclColumnName("ROWID");
          tabCol.setExcelColumnNo(rowIDColNum);
          super.getColumnMapping().addColumn(tabCol);
        }
*/        
    }
    
}
