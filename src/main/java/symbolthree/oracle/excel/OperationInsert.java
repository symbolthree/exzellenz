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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/OperationInsert.java $
 * $Author: Christopher Ho $
 * $Date: 7/14/16 9:44p $
 * $Revision: 16 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import oracle.jdbc.OracleCallableStatement;
import oracle.jdbc.OracleTypes;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.sql.*;
import java.util.*;

public class OperationInsert extends Operation implements Constants {

    private int SHOWING_ROWCOUNT = EXZProp.instance().getInt(EXZ_LOG_INTERVAL);
    
    public OperationInsert() {
        super();
    }

    @Override
    public void doOperation() throws EXZException {
        String tableName  = EXZParams.instance().getValue(TABLE_NAME);
        String objectType = super.getObjectType();

        if (!objectType.equals("TABLE")) {
            throw new EXZException(EXZI18N.inst().get("ERR.OPERATION_MODE",
                    EXZParams.instance().getValue(OPERATION_MODE), tableName));
        }

        Workbook wb = super.getWorkbook();

        try {
            Connection conn = super.getConnection();

            // create SQL Statement
            ArrayList<Integer>    colPosition   = new ArrayList<Integer>();
            ArrayList<String>     colType       = new ArrayList<String>();
            String                prepareSQL    = "INSERT INTO " + tableName + "(";
            String                questionMarks = "";
            Iterator<TableColumn> itr           = super.getColumnMapping().getColumns().iterator();
            int                   noDataCols    = 0;

            while (itr.hasNext()) {
                TableColumn tabCol = (TableColumn) itr.next();

                String columnName = tabCol.getColumnName();
                
                if (tabCol.isNeeded() && ! tabCol.isResultColumn() && ! tabCol.isRowIDColumn()) {
                  
                  if (tabCol.getExcelColumnNo() == 0) {
                  
                    EXZHelper.log(LOG_INFO, columnName + " is not available");

                  } else {

                    colPosition.add(new Integer(tabCol.getExcelColumnNo()));
                    colType.add(tabCol.getColumnType());
                    prepareSQL    = prepareSQL + columnName + ",";
                    questionMarks = questionMarks + "?,";
                    noDataCols++;
                    
                  }
                }
            }

            prepareSQL    = prepareSQL.substring(0, prepareSQL.length() - 1);
            questionMarks = questionMarks.substring(0, questionMarks.length() - 1);
            prepareSQL    = prepareSQL + ") VALUES (" + questionMarks + ")";
            
            // add rowid result
            prepareSQL = prepareSQL + " RETURNING ROWIDTOCHAR(ROWID) INTO ?";
            
            EXZHelper.log(LOG_DEBUG, prepareSQL);

            Sheet dataSheet = super.getWorkbook().getSheet(EXZParams.instance().getValue(DATA_WORKSHEET));
            
            //PreparedStatement preStmt         = conn.prepareStatement(prepareSQL);
            OracleCallableStatement preStmt = (OracleCallableStatement)conn.prepareCall(prepareSQL);
            
            int    rowSuccess      = 0;
            int    rowFailure      = 0;
            int    rowSkipped      = 0;
            int    lastRowNumber   = dataSheet.getLastRowNum();
            int    dataRowNumber   = EXZParams.instance().getInt(COLUMN_TITLE_ROW) + 1;
            int    resultColNumber = super.getColumnMapping().getResultColumn();
            int    ROWIDColNumber  = super.getColumnMapping().getROWIDColumn();
            String ROWID           = "";
            
            EXZHelper.log(LOG_DEBUG, "Last Row Number = " + (lastRowNumber + 1));
            EXZHelper.log(LOG_DEBUG, "resultColNumber = " + resultColNumber);
            EXZHelper.log(LOG_DEBUG, "ROWIDColNumber  = " + ROWIDColNumber);

            if (resultColNumber == 0) {
                throw new EXZException("Cannot find result column ["
                                       + EXZParams.instance().getValue(RESULT_COLUMN_NAME) + "]");
            }

            //EXZHelper.log(LOG_DEBUG, "resultColNumber:" + EXZHelper.number2Letter(resultColNumber));

            String    processed    = null;
            CellStyle successStyle = (super.getSpecialCells().get(RESULT_SUCCESS)).getCellStyle();
            CellStyle failureStyle = (super.getSpecialCells().get(RESULT_FAILURE)).getCellStyle();

            EXZHelper.log(LOG_DEBUG, "Process start...");

            while (dataRowNumber <= lastRowNumber + 1) {

                // skip row if has flag for processed
                processed = EXZHelper.readString(wb, dataSheet, dataRowNumber, resultColNumber);

                // The lastRowNumber sometimes is not really the last row,
                // so we need to break the loop when the entire row (of all data columns) is empty
                boolean isAllEmpty   = true;
                boolean isErrorFound = false;

                if (!processed.equals(EXZParams.instance().getValue(RESULT_SUCCESS)) &&
                	!processed.equals(EXZParams.instance().getValue(RESULT_FAILURE)) &&
                	!processed.equals(EXZParams.instance().getValue(PENDING_UPDATE))) {
                    try {
                        for (int i = 0; i < colPosition.size(); i++) {
                            if (colType.get(i).equals("VARCHAR2")) {
                                String value = EXZHelper.readString(wb, dataSheet, dataRowNumber,
                                                   toInt(colPosition.get(i)));

                                isAllEmpty = value.equals("") && isAllEmpty;
                                preStmt.setString(i + 1, value);
                                
                            } else if (colType.get(i).equals("NUMBER")) {
                                double value = EXZHelper.readDouble(wb, dataSheet, dataRowNumber,
                                                   toInt(colPosition.get(i)));

                                isAllEmpty = (value == Double.MIN_VALUE) && isAllEmpty;

                                if (value == Double.MIN_VALUE) {
                                    preStmt.setObject(i + 1, null);
                                } else {
                                    preStmt.setDouble(i + 1, value);
                                }
                                
                            } else if (colType.get(i).equals("INTEGER")) {
                                int value = toInt(EXZHelper.readString(wb, dataSheet, dataRowNumber,
                                                toInt(colPosition.get(i))));

                                isAllEmpty = (value == Integer.MIN_VALUE) && isAllEmpty;

                                if (value == Integer.MIN_VALUE) {
                                    preStmt.setObject(i + 1, null);
                                } else {
                                    preStmt.setInt(i + 1, value);
                                }
                                
                            } else if (colType.get(i).equals("DATE")) {
                                java.sql.Date value = EXZHelper.readDate(wb, dataSheet, dataRowNumber,
                                                          toInt(colPosition.get(i)));

                                isAllEmpty = (value == null) && isAllEmpty;
                                preStmt.setDate(i + 1, value);
                            } else {
                                String value = EXZHelper.readString(wb, dataSheet, dataRowNumber,
                                                   toInt(colPosition.get(i)));

                                isAllEmpty = value.equals("") && isAllEmpty;
                                preStmt.setObject(i + 1, value);
                            }    // end type switch (if-then-else)
                        }        // end row scanning (for loop)

                        if (isAllEmpty) {
                            EXZHelper.log(LOG_INFO, "Row " + dataRowNumber + " is empty. Process stopped.");

                            break;
                        } else {
                            EXZHelper.log(LOG_DEBUG, "Row " + dataRowNumber + " is not empty.");
                        }

                        //preStmt.execute();
                        // get ROWID back
                        preStmt.registerReturnParameter(noDataCols+1, OracleTypes.VARCHAR);
                        preStmt.executeUpdate();
            		    ResultSet rs = preStmt.getReturnResultSet();
            		    rs.next();
            		    ROWID = rs.getString(1);
                        
            		    Cell cell = null;
                        if (wb instanceof HSSFWorkbook) {
                            cell = (HSSFCell) cell;
                        } else if (wb instanceof XSSFWorkbook) {
                            cell = (XSSFCell) cell;
                        }
                        
                        // skip rowid column since 1.11
                        if (createROWIDFlag()) { 
	                        Row row = dataSheet.getRow(dataRowNumber - 1);
	                        cell = row.createCell(ROWIDColNumber - 1);
	                        if (wb instanceof HSSFWorkbook) {                        
	                          cell.setCellValue(new HSSFRichTextString(ROWID));
	                        } else if (wb instanceof XSSFWorkbook) {
	                          cell.setCellValue(new XSSFRichTextString(ROWID));
	                        }
                        }
            		    		
                    } catch (Exception ee) {
                    	
                        Cell cell = dataSheet.getRow(dataRowNumber - 1).createCell(resultColNumber - 1);

                        EXZHelper.log(LOG_ERROR, "Error for data in row " + dataRowNumber);
                        EXZHelper.logError(ee);

                        String errMsg = ee.getMessage();
                        if (errMsg == null) {
                        	errMsg = EXZParams.instance().getValue(RESULT_FAILURE);
                        }
                        
                        if (cell instanceof HSSFCell) {
                            cell.setCellStyle((HSSFCellStyle) failureStyle);
                            cell.setCellValue(new HSSFRichTextString(errMsg));
                        } else if (cell instanceof XSSFCell) {
                            cell.setCellStyle((XSSFCellStyle) failureStyle);
                            cell.setCellValue(new XSSFRichTextString(errMsg));
                        }

                        rowFailure++;

                        // error handling cases
                        if (EXZParams.instance().getValue(ERROR_HANDLING).equals(COMMIT_AND_EXIT)) {
                        	// commit and exit
                            conn.commit();
                            super.writeWorkbook();

                            throw new EXZException(ee);
                        }

                        if (EXZParams.instance().getValue(ERROR_HANDLING).equals(NO_COMMIT_AND_EXIT)) {
                        	// no commit and exit
                            super.writeWorkbook();
                            throw new EXZException(ee);
                        }

                        if (EXZParams.instance().getValue(ERROR_HANDLING).equals(CONTINUE_ON_ERROR)) {
                            isErrorFound = true;
                        }
                    }

                    if (!isErrorFound) {
                        EXZHelper.log(LOG_DEBUG, "Row " + dataRowNumber + " successfully processed");

                        Cell cell = dataSheet.getRow(dataRowNumber - 1).createCell(resultColNumber - 1);
                        
                        //set success column
                        if (cell instanceof HSSFCell) {
                            cell.setCellStyle((HSSFCellStyle) successStyle);
                            cell.setCellValue(new HSSFRichTextString(EXZParams.instance().getValue(RESULT_SUCCESS)));
                        } else if (cell instanceof XSSFCell) {
                            cell.setCellStyle((XSSFCellStyle) successStyle);
                            cell.setCellValue(new XSSFRichTextString(EXZParams.instance().getValue(RESULT_SUCCESS)));
                        }

                        rowSuccess++;
                        
                        if (rowSuccess%SHOWING_ROWCOUNT == 0) {
                          EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.INSERT_ROW_DATA", String.valueOf(rowSuccess)));
                        }                              
                    }

                // if the row has been flagged to skip
                } else {
                    EXZHelper.log(LOG_DEBUG, "Data in row " + dataRowNumber + " is skipped (" + processed + ")");
                    rowSkipped++;
                }

                dataRowNumber++;
            }    // end while-loop for row scanning

            conn.commit();
            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.ROW_PROCESSED", String.valueOf(rowSuccess)));
            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.ROW_FAILED", String.valueOf(rowFailure)));
            EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.ROW_SKIPPED", String.valueOf(rowSkipped)));

            // hide ROWID column
            if (createROWIDFlag()) {
	            if (objectType.equals("TABLE")) {
	                int rowidColNo = super.getColumnMapping().getROWIDColumn();
	
	                if (!dataSheet.isColumnHidden(rowidColNo - 1)) {
	                    dataSheet.setColumnHidden(rowidColNo - 1, true);
	                }
	            }
            }
            
            // check Excel file is required to be saved
            if ((rowSuccess > 0) || (rowFailure > 0)) {
                super.setSaveFileRequired(true);
            } else {
                super.setSaveFileRequired(false);
                EXZHelper.log(LOG_DEBUG, "No need to save " + super.getExcelFile().getAbsolutePath());
            }
        } catch (Exception e) {
            throw new EXZException(e);
        }
    }

    @Override
    public void defineTableName() throws EXZException {
        super.defineTableName();
        if (createROWIDFlag()) {
          super.addROWIDCol();
        }
    }
}
