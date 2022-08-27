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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/TableColumn.java $
 * $Author: Christopher Ho $
 * $Date: 7/12/16 11:09a $
 * $Revision: 7 $
******************************************************************************/

package symbolthree.oracle.excel;

public class TableColumn {
    private boolean resultColumn = false;
    private boolean needed       = true;
    private boolean rowIDColumn  = false;
    private String  columnName;
    private boolean columnNullable;
    private int     columnSize;
    private String  columnType;
    private int     excelColumnNo;
    private String  execlColumnName;
    private boolean nameMatched;

    /**
     * @param columnName the columnName to set
     */
    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    /**
     * @return the columnName
     */
    public String getColumnName() {
        return columnName;
    }

    /**
     * @param columnSize the columnSize to set
     */
    public void setColumnSize(int columnSize) {
        this.columnSize = columnSize;
    }

    /**
     * @return the columnSize
     */
    public int getColumnSize() {
        return columnSize;
    }

    /**
     * @param columnType the columnType to set
     */
    public void setColumnType(String columnType) {
        this.columnType = columnType;
    }

    /**
     * @return the columnType
     */
    public String getColumnType() {
        return columnType;
    }

    /**
     * @param columnNullable the columnNullable to set
     */
    public void setColumnNullable(boolean columnNullable) {
        this.columnNullable = columnNullable;
    }

    /**
     * @return the columnNullable
     */
    public boolean isColumnNullable() {
        return columnNullable;
    }

    /**
     * @param execlColumnName the execlColumnName to set
     */
    public void setExeclColumnName(String execlColumnName) {
        this.execlColumnName = execlColumnName;
    }

    /**
     * @return the execlColumnName
     */
    public String getExeclColumnName() {
        return execlColumnName;
    }

    /**
     * @param nameMatched the nameMatched to set
     */
    public void setNameMatched(boolean nameMatched) {
        this.nameMatched = nameMatched;
    }

    /**
     * @return the nameMatched
     */
    public boolean isNameMatched() {
        return nameMatched;
    }

    /**
     * @param excelColumnNo the excelColumnNo to set
     */
    public void setExcelColumnNo(int excelColumnNo) {
        this.excelColumnNo = excelColumnNo;
    }

    /**
     * @return the excelColumnNo
     */
    public int getExcelColumnNo() {
        return excelColumnNo;
    }

    /**
     * @param needed the needed to set
     */
    public void setNeeded(boolean needed) {
        this.needed = needed;
    }

    /**
     * @return the needed
     */
    public boolean isNeeded() {
        return needed;
    }

    /**
     * @param resultColumn the resultColumn to set
     */
    public void setResultColumn(boolean resultColumn) {
        this.resultColumn = resultColumn;
    }

    /**
     * @return the resultColumn
     */
    public boolean isResultColumn() {
        return resultColumn;
    }

    /**
     * @param rowIDColumn the rowIDColumn to set
     */
    public void setRowIDColumn(boolean rowIDColumn) {
        this.rowIDColumn = rowIDColumn;
    }

    /**
     * @return the rowIDColumn
     */
    public boolean isRowIDColumn() {
        return rowIDColumn;
    }
}
