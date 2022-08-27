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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZI18N.java $
 * $Author: Christopher Ho $
 * $Date: 7/12/16 11:09a $
 * $Revision: 4 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.commons.i18n.MessageManager;
import org.apache.commons.i18n.XMLMessageProvider;

//~--- JDK imports ------------------------------------------------------------

import java.io.InputStream;

import java.util.Locale;
import java.util.StringTokenizer;

public class EXZI18N implements Constants {
    private static EXZI18N exzI18N = null;
    private Locale         locale;

    public EXZI18N() {
        try {
            InputStream fis = this.getClass().getResourceAsStream("/symbolthree/oracle/excel/EXZ_STRING.xml");

            XMLMessageProvider.install("string", fis);

            String lang = EXZProp.instance().getStr(PROGRAM_LOCALE);

            // use English if runing in command line
            if (System.getProperty(RUN_MODE).equals("GUI")) {
                if (lang.equals("ZHT")) {
                    locale = new Locale("zh", "TW");
                } else if (lang.equals("US")) {
                    locale = new Locale("en", "US");
                } else {
                    locale = new Locale("en", "US");
                }
            } else {
                locale = new Locale("en", "US");
            }

            EXZHelper.log(EXZHelper.LOG_DEBUG, "Locale=" + locale.getDisplayName());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static EXZI18N inst() {
        if (exzI18N == null) {
            exzI18N = new EXZI18N();
        }

        return exzI18N;
    }

    private String get(String msgKey, Object[] subStrs) {
        try {
            StringTokenizer st   = new StringTokenizer(msgKey, ".");
            String          name = st.nextToken();
            String          key  = st.nextToken();

            if ((msgKey == null) || msgKey.equals("")) {
                return "";
            }

            String rtnStr = MessageManager.getText(name, key, subStrs, locale);

            if ((rtnStr == null) || rtnStr.equals("")) {
                rtnStr = MessageManager.getText(name, key, subStrs, Locale.US);
            }

            if ((rtnStr == null) || rtnStr.equals("")) {
                rtnStr = msgKey;
            }
            
            rtnStr = rtnStr.replaceAll("(\\\\n)", System.getProperty("line.separator"));
            
            return rtnStr;
        } catch (Exception e) {
            return msgKey;
        }
    }

    public String get(String msgKey, String subStr1, String subStr2) {
        return get(msgKey, new Object[] { subStr1, subStr2 });
    }

    public String get(String msgKey, String substr1) {
        return get(msgKey, new Object[] { substr1 });
    }

    public String get(String msgKey) {
        return get(msgKey, new Object[0]);
    }
}
