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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZParams.java $
 * $Author: Christopher Ho $
 * $Date: 2/17/17 9:58a $
 * $Revision: 9 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- JDK imports ------------------------------------------------------------

import java.util.*;

public class EXZParams implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZParams.java 9     2/17/17 9:58a Christopher Ho $";
    private static EXZParams exzParams = null;
    private Properties       prop      = null;

    protected EXZParams() {
        prop = new Properties();
    }

    public static EXZParams instance() {
        if (exzParams == null) {
            exzParams = new EXZParams();
        }

        return exzParams;
    }

    public void setValue(String param, String value) {
        if (!EXZHelper.isEmpty(value)) {
            prop.setProperty(param.toUpperCase(), value);
        }
    }

    public String getJDBUrl() {
    	/*
        return "jdbc:oracle:thin:@" + prop.getProperty(SERVER) + ":" + prop.getProperty(PORT) + ":"
               + prop.getProperty(SID);
        */
    	return prop.getProperty(JDBC_URL);
    }

    /*
    public String getJDBCUrlWithCredentials() {
   	
        return "jdbc:oracle:thin:" + prop.getProperty(USERNAME) + "/" + prop.getProperty(PASSWORD) + "@"
               + prop.getProperty(SERVER) + ":" + prop.getProperty(PORT) + ":" + prop.getProperty(SID);
    }
    */    

    public String getValue(String param) {
        return prop.getProperty(param);
    }

    public int getInt(String param) {
        return Integer.parseInt(prop.getProperty(param));
    }
    
    public double getDouble(String param) {
    	return Double.parseDouble(prop.getProperty(param));
        //return Integer.parseInt(prop.getProperty(param));
    }
    
    public void clear() {
    	prop.clear();
    }
    
}
