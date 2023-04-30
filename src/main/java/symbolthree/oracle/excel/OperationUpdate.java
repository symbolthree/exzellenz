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

//~--- JDK imports ------------------------------------------------------------

import java.sql.*;

import java.util.*;

public class OperationUpdate extends Operation {

    public OperationUpdate() {
        super();
    }

    @Override
    public void doOperation() throws EXZException {
        String tableName  = EXZParams.instance().getValue(TABLE_NAME);
        String owner      = EXZParams.instance().getValue(OWNER);
        if (owner != null && !owner.equals("")) {
        	tableName = owner + "." + tableName;
        }        
        String objectType = super.getObjectType();

        if (!objectType.equals("TABLE")) {
            throw new EXZException(EXZI18N.inst().get("ERR.OPERATION_MODE",
                    EXZParams.instance().getValue(OPERATION_MODE), tableName));
        }

        Workbook wb = super.getWorkbook();

        try {
            Connection conn = super.getConnection();

            // create SQL Statement
            ArrayList<Integer>    colPosition = new ArrayList<Integer>();
            ArrayList<String>     colType     = new ArrayList<String>();
            String                prepareSQL  = "UPDATE " + EXZParams.instance().getValue(TABLE_NAME) + " SET ";
            Iterator<TableColumn> itr         = super.getColumnMapping().getColumns().iterator();
            int                   rowIDPos    = 0;

            while (itr.hasNext()) {
                TableColumn tabCol = (TableColumn) itr.next();

/*
                if (tabCol.isNeeded() &&!tabCol.isResultColumn() &&!tabCol.getColumnType().equals("ROWID")) {
                    String columnName = tabCol.getColumnName();

                    colPosition.add(new Integer(tabCol.getExcelColumnNo()));
                    colType.add(tabCol.getColumnType());
                    prepareSQL = prepareSQL + columnName + "=?,";
                }

                if (tabCol.isNeeded() &&!tabCol.isResultColumn() && tabCol.getColumnType().equals("ROWID")) {
                    rowIDPos = tabCol.getExcelColumnNo();
                }
*/
                if ((tabCol.getExcelColumnNo() > 0) &&!tabCol.isResultColumn()
                        &&!tabCol.getColumnType().equals("ROWID")) {
                    String columnName = tabCol.getColumnName();

                    colPosition.add(Integer.valueOf(tabCol.getExcelColumnNo()));
                    colType.add(tabCol.getColumnType());
                    prepareSQL = prepareSQL + columnName + "=?,";
                }

                if ((tabCol.getExcelColumnNo() > 0) &&!tabCol.isResultColumn()
                        && tabCol.getColumnType().equals("ROWID")) {
                    rowIDPos = tabCol.getExcelColumnNo();
                }
            }

            colPosition.add(Integer.valueOf(rowIDPos));
            colType.add("ROWID");
            prepareSQL = prepareSQL.substring(0, prepareSQL.length() - 1);
            prepareSQL = prepareSQL + " WHERE ROWID=?";
            EXZHelper.log(LOG_DEBUG, prepareSQL);

            String dataSheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
            Sheet  dataSheet     = super.getWorkbook().getSheet(dataSheetName);

            dataSheet.setSelected(true);

            PreparedStatement preStmt         = conn.prepareStatement(prepareSQL);
            int               rowSuccess      = 0;
            int               rowFailure      = 0;
            int               rowSkipped      = 0;
            int               lastRowNumber   = dataSheet.getLastRowNum();
            int               dataRowNumber   = EXZParams.instance().getInt(COLUMN_TITLE_ROW) + 1;
            int               resultColNumber = super.getColumnMapping().getResultColumn();

            EXZHelper.log(LOG_DEBUG, "Last Row Number=" + (lastRowNumber + 1));

            if (resultColNumber == 0) {
                throw new EXZException("Cannot find result column ["
                                       + EXZParams.instance().getValue(RESULT_COLUMN_NAME) + "]");
            }

            EXZHelper.log(LOG_DEBUG, "resultColNumber:" + EXZHelper.number2Letter(resultColNumber));

            String    processed    = null;
            CellStyle successStyle = super.getSpecialCells().get(RESULT_SUCCESS).getCellStyle();
            CellStyle failureStyle = super.getSpecialCells().get(RESULT_FAILURE).getCellStyle();

            EXZHelper.log(LOG_DEBUG, "PENDING_UPDATE STRING = " + EXZParams.instance().getValue(PENDING_UPDATE));
            EXZHelper.log(LOG_DEBUG, "Process start...");

            while (dataRowNumber <= lastRowNumber + 1) {

                // skip row if has flag for processed
                processed = EXZHelper.readString(wb, dataSheet, dataRowNumber, resultColNumber);

                // The lastRowNumber sometimes is not really the last row,
                // so we need to break the loop when the entire row (of all data columns) is empty
                boolean isAllEmpty   = true;
                boolean isErrorFound = false;

                if (processed.equals(EXZParams.instance().getValue(PENDING_UPDATE))) {
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
                            } else if (colType.get(i).equals("ROWID")) {
                                String value = EXZHelper.readString(wb, dataSheet, dataRowNumber,
                                                   toInt(colPosition.get(i)));

                                isAllEmpty = value.equals("") && isAllEmpty;
                                preStmt.setString(i + 1, value);
                                // DEBUG 
                                EXZHelper.log(LOG_DEBUG, "ROWID = " + value);
                                
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
                        }

                        preStmt.executeUpdate();

                    } catch (Exception ee) {
                        Cell cell = dataSheet.getRow(dataRowNumber - 1).createCell(resultColNumber - 1);

                        EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.ROW_DATA", String.valueOf(dataRowNumber)));
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
                        EXZHelper.log(LOG_DEBUG, "ERROR_HANDLING=" + EXZParams.instance().getValue(ERROR_HANDLING));

                        // error handling cases
                        if (EXZParams.instance().getValue(ERROR_HANDLING).equals(COMMIT_AND_EXIT)) {
                            conn.commit();
                            super.writeWorkbook();

                            throw new EXZException(ee);
                        }

                        if (EXZParams.instance().getValue(ERROR_HANDLING).equals(NO_COMMIT_AND_EXIT)) {

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

                        if (cell instanceof HSSFCell) {
                            cell.setCellStyle((HSSFCellStyle) successStyle);
                            cell.setCellValue(new HSSFRichTextString(EXZParams.instance().getValue(RESULT_SUCCESS)));
                        } else if (cell instanceof XSSFCell) {
                            cell.setCellStyle((XSSFCellStyle) successStyle);
                            cell.setCellValue(new XSSFRichTextString(EXZParams.instance().getValue(RESULT_SUCCESS)));
                        }

                        rowSuccess++;
                    }

                    // if the row has been flagged for processed...
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

            // check Excel file is required to be saved
            if ((rowSuccess > 0) || (rowFailure > 0)) {
                wb.setFirstVisibleTab(wb.getSheetIndex(dataSheetName));
                super.setSaveFileRequired(true);
            } else {
                super.setSaveFileRequired(false);
                EXZHelper.log(LOG_INFO, "No need to save " + super.getExcelFile().getAbsolutePath());
            }
        } catch (Exception e) {
            throw new EXZException(e);
        }
    }

    @Override
    public void defineTableName() throws EXZException {
        super.defineTableName();

        // add ROWID column and check existence
        TableColumn rowID = new TableColumn();

        rowID.setColumnName("ROWID");
        rowID.setExcelColumnName("ROWID");

        Workbook wb            = super.getWorkbook();
        String   worksheetName = EXZParams.instance().getValue(DATA_WORKSHEET);
        Sheet    sheet         = wb.getSheet(worksheetName);
        int      rowNo         = EXZParams.instance().getInt(COLUMN_TITLE_ROW);
        Row      row           = sheet.getRow(rowNo - 1);
        int      lastColNo     = row.getLastCellNum();
        int      colNo         = 1;

        while (colNo <= lastColNo) {
            String colName = EXZHelper.readString(wb, sheet, rowNo, colNo);

            if (colName.equals("ROWID")) {
                rowID.setExcelColumnNo(colNo);
                rowID.setNameMatched(true);
                rowID.setColumnType("ROWID");
                rowID.setRowIDColumn(true);

                break;
            }

            colNo++;
        }

        if (rowID.getExcelColumnNo() == 0) {
            EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.NO_ROWID", worksheetName));

            throw new EXZException();
        } else {
            super.getColumnMapping().addColumn(rowID);
        }
    }

    @Override
    public void createColumnMapping(Map<String, String> customColumnMap) throws EXZException, SQLException {
        super.createColumnMapping(customColumnMap);
    }
}
