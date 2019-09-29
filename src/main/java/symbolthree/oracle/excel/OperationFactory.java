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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/OperationFactory.java $
 * $Author: Christopher Ho $
 * $Date: 7/12/16 11:09a $
 * $Revision: 6 $
******************************************************************************/

package symbolthree.oracle.excel;

public class OperationFactory implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/OperationFactory.java 6     7/12/16 11:09a Christopher Ho $";
    private String opMode = null;

    public OperationFactory(String mode) {
        opMode = mode.substring(0, 1).toUpperCase() + mode.substring(1, mode.length()).toLowerCase();
    }

    public Operation getOperation() throws EXZException {
        Operation op        = null;
        String    className = OPERATION_CLASS_SUFFIX + opMode;

        try {
            Class<?> clazz = Class.forName(className);

            op = (Operation) clazz.newInstance();

            return op;
        } catch (ClassNotFoundException cnfe) {
            EXZHelper.log(LOG_ERROR, "Cannot find class " + className);
            EXZHelper.log(LOG_ERROR, "Please check your operation mode parameter.");
            EXZHelper.logError(cnfe);

            throw new EXZException(cnfe);
        } catch (IllegalAccessException iae) {
            EXZHelper.log(LOG_ERROR, "Illegal access of class " + className);
            EXZHelper.logError(iae);

            throw new EXZException(iae);
        } catch (InstantiationException ie) {
            EXZHelper.log(LOG_ERROR, "Cannot create instance of class " + className);
            EXZHelper.logError(ie);

            throw new EXZException(ie);
        }
    }
}
