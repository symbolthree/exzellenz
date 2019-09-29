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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/ColumnMapping.java $
 * $Author: Christopher Ho $
 * $Date: 7/12/16 11:09a $
 * $Revision: 10 $
******************************************************************************/



package symbolthree.oracle.excel;

//~--- JDK imports ------------------------------------------------------------

import java.util.*;

public class ColumnMapping implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/ColumnMapping.java 10    7/12/16 11:09a Christopher Ho $";
    private ArrayList<TableColumn> tableColumns = new ArrayList<TableColumn>();
    private String                 objectType;

    public ArrayList<TableColumn> getColumns() {
        return tableColumns;
    }

    public void addColumn(TableColumn _tabCol) {

        // check whether this column has custom column name
        if (tableColumns.isEmpty()) {
            tableColumns.add(_tabCol);
        } else {
            boolean               colExist = false;
            Iterator<TableColumn> itr      = tableColumns.iterator();

            while (itr.hasNext() &&!colExist) {
                TableColumn tabCol = (TableColumn) itr.next();

                if (!tabCol.isResultColumn()) {
                    if (tabCol.getColumnName().equals(_tabCol.getColumnName())) {
                        tabCol.setColumnType(_tabCol.getColumnType());
                        tabCol.setColumnSize(_tabCol.getColumnSize());
                        tabCol.setColumnNullable(_tabCol.isColumnNullable());
                        tabCol.setNeeded(_tabCol.isNeeded());
                        tabCol.setNameMatched(true);
                        colExist = true;

                        break;
                    }
                }
            }

            if (!colExist) {
                tableColumns.add(_tabCol);
            }
        }
    }

    public boolean checkMapping() {
        boolean               rtnValue = true;
        Iterator<TableColumn> itr      = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            EXZHelper.log(LOG_DEBUG, "Checking [" + tabCol.getColumnName() + ", " + tabCol.getExeclColumnName() + "]");
            
            if (!tabCol.isResultColumn() &&
            	!tabCol.isRowIDColumn()) {
            	
                if (! tabCol.isNameMatched() &&
                	! EXZHelper.isEmpty(tabCol.getExeclColumnName())) {
                	
                    rtnValue = false;
                    EXZHelper.log(LOG_ERROR,
                                  EXZI18N.inst().get("ERR.INVALID_MAPPING", tabCol.getColumnName(),
                                                     tabCol.getExeclColumnName()));
                }

                if (! tabCol.isNameMatched() && 
                	EXZHelper.isEmpty(tabCol.getExeclColumnName())) {
                    
                	tabCol.setExeclColumnName(tabCol.getColumnName());
                }
            }
        }

        return rtnValue;
    }

    public void showMapping() {
        String tab = "\t";

        EXZHelper.log(LOG_DEBUG,
                      "Column Name" + tab + "Column Type" + tab + "Excel Column Name" + tab + "Excel Col Pos" + tab
                      + "Nullable" + tab + "isMatched" + tab + "isNeeded" + tab + "isResult");

        Iterator<TableColumn> itr = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            EXZHelper.log(LOG_DEBUG,
                          tabCol.getColumnName() + tab + tabCol.getColumnType() + tab + tabCol.getExeclColumnName()
                          + tab + tabCol.getExcelColumnNo() + tab + tabCol.isColumnNullable() + tab
                          + tabCol.isNameMatched() + tab + tabCol.isNeeded() + tab + tabCol.isResultColumn());
        }
    }

    public int getResultColumn() {
        int                   rtnValue = 0;
        Iterator<TableColumn> itr      = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            if (tabCol.isResultColumn()) {
                rtnValue = tabCol.getExcelColumnNo();

                break;
            }
        }

        return rtnValue;
    }

    public int getROWIDColumn() {
        int                   rtnValue = 0;
        Iterator<TableColumn> itr      = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            if (tabCol.getColumnName() != null && tabCol.getColumnName().equals("ROWID")) {
                rtnValue = tabCol.getExcelColumnNo();

                break;
            }
        }

        return rtnValue;
    }

    public void setExcelColumnNo(String columnName, int excelColNo) {
        Iterator<TableColumn> itr = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            if (tabCol.getExeclColumnName().equals(columnName)) {
                tabCol.setExcelColumnNo(excelColNo);

                break;
            }
        }
    }

    public int getMaxExcelColumnNo() {
        int                   maxColNo = 0;
        Iterator<TableColumn> itr      = tableColumns.iterator();

        while (itr.hasNext()) {
            TableColumn tabCol = (TableColumn) itr.next();

            if (maxColNo < tabCol.getExcelColumnNo()) {
                maxColNo = tabCol.getExcelColumnNo();
            }
        }

        return maxColNo;
    }

    public void setObjectType(String objectType) {
        this.objectType = objectType;
    }

    public String getObjectType() {
        return objectType;
    }
}
