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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZProp.java $
 * $Author: Christopher Ho $
 * $Date: 7/14/16 9:44p $
 * $Revision: 16 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- JDK imports ------------------------------------------------------------

import java.io.*;
import java.util.*;

import org.apache.commons.io.FileUtils;

public class EXZProp implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZProp.java 16    7/14/16 9:44p Christopher Ho $";
    private static final String RESOURCE_BUNDLE = "EXZ.properties";
    private static EXZProp      exzInstance     = null;
    private Properties          prop;

    public EXZProp() {
        prop = new Properties();

        try {
            
        	File f = new File(EXZ_APPLICATION_DIR);
        	if (! f.exists()) FileUtils.forceMkdir(f);

        	File propUser     = new File(EXZ_APPLICATION_DIR, RESOURCE_BUNDLE);
        	if (! propUser.exists()) {
        	  File propTemplate = new File(System.getProperty("user.dir"), RESOURCE_BUNDLE);
        	  FileUtils.copyFile(propTemplate, propUser);
        	}
        	
            FileInputStream is = new FileInputStream(propUser);
            prop.load(is);
            is.close();
            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static EXZProp instance() {
        if (exzInstance == null) {
            exzInstance = new EXZProp();
        }

        return exzInstance;
    }

    public int getInt(String key) {
        String _str = prop.getProperty(key);

        if (_str == null) {
            return 0;
        } else {
            return Integer.valueOf(_str).intValue();
        }
    }

    public String getStr(String key) {
        String _str = prop.getProperty(key);

        if (_str == null) {
            return null;
        } else {
            return _str;
        }
    }

    public boolean getBoolean(String key) {
        return Boolean.valueOf(prop.getProperty(key)).booleanValue();
    }

    public Properties getSessionProp() {
        Properties          sessionProp = new Properties();
        Enumeration<Object> key         = prop.keys();

        while (key.hasMoreElements()) {
            String name = (String) key.nextElement();

            if (name.startsWith("v$session")) {
                String value = prop.getProperty(name);

                EXZHelper.log(LOG_DEBUG, name + "," + value);
                sessionProp.setProperty(name, value);
            }
        }

        return sessionProp;
    }
    
    public void setStr(String key, String value) {
      prop.setProperty(key, value);
    }
    
    public void saveSettings() {
      try {
        File           tempfile = File.createTempFile("exz", "tmp");
        FileWriter     fw       = new FileWriter(tempfile);
        File           propfile = new File(EXZ_APPLICATION_DIR, RESOURCE_BUNDLE);
        BufferedReader reader   = new BufferedReader(new FileReader(propfile));
        String         line     = null;

        while ((line = reader.readLine()) != null) {
          if (line.startsWith(LAST_FILE)) {
              line = LAST_FILE + "=" + propStr(prop.getProperty(LAST_FILE));
          }
          /*
          if (line.startsWith(SAVE_NEW_FILE)) {
            line = SAVE_NEW_FILE + "=" + prop.getProperty(SAVE_NEW_FILE);
          }
          */
          if (line.startsWith(EXZ_LOG_LEVEL)) {
        	  line = EXZ_LOG_LEVEL + "=" + prop.getProperty(EXZ_LOG_LEVEL);
          }
          if (line.startsWith(EXZ_LOG_INTERVAL)) {
        	  line = EXZ_LOG_INTERVAL + "=" + prop.getProperty(EXZ_LOG_INTERVAL);
          }
          if (line.startsWith(WINDOW_WIDTH)) {
        	  line = WINDOW_WIDTH + "=" + prop.getProperty(WINDOW_WIDTH);
          }
          if (line.startsWith(WINDOW_HEIGHT)) {
        	  line = WINDOW_HEIGHT + "=" + prop.getProperty(WINDOW_HEIGHT);
          }
          if (!line.startsWith("#~")) {
              //fw.write(line + System.lineSeparator());
              fw.write(line + System.getProperty("line.separator"));
          }
        }
        fw.flush();
        fw.close();
        reader.close();

        propfile.delete();
        FileUtils.moveFile(tempfile, propfile);

    } catch (Exception e) {
        EXZHelper.logError(e);
    }
  }
    
    private static String propStr(String theString) {
      return propStr(theString, true);
    }
    
    private static String propStr(String theString, boolean escapeSpace) {
      int len    = theString.length();
      int bufLen = len * 2;

      if (bufLen < 0) {
          bufLen = Integer.MAX_VALUE;
      }

      StringBuffer outBuffer = new StringBuffer(bufLen);

      for (int x = 0; x < len; x++) {
          char aChar = theString.charAt(x);

          // Handle common case first, selecting largest block that
          // avoids the specials below
          if ((aChar > 61) && (aChar < 127)) {
              if (aChar == '\\') {
                  outBuffer.append('\\');
                  outBuffer.append('\\');

                  continue;
              }

              outBuffer.append(aChar);

              continue;
          }

          switch (aChar) {
          case ' ' :
              if ((x == 0) || escapeSpace) {
                  outBuffer.append('\\');
              }

              outBuffer.append(' ');

              break;

          case '\t' :
              outBuffer.append('\\');
              outBuffer.append('t');

              break;

          case '\n' :
              outBuffer.append('\\');
              outBuffer.append('n');

              break;

          case '\r' :
              outBuffer.append('\\');
              outBuffer.append('r');

              break;

          case '\f' :
              outBuffer.append('\\');
              outBuffer.append('f');

              break;

          case '=' :    // Fall through
          case ':' :    // Fall through
          case '#' :    // Fall through
          case '!' :
              outBuffer.append('\\');
              outBuffer.append(aChar);

              break;

          default :
              if ((aChar < 0x0020) || (aChar > 0x007e)) {
                  outBuffer.append('\\');
                  outBuffer.append('u');
                  outBuffer.append(toHex((aChar >> 12) & 0xF));
                  outBuffer.append(toHex((aChar >> 8) & 0xF));
                  outBuffer.append(toHex((aChar >> 4) & 0xF));
                  outBuffer.append(toHex(aChar & 0xF));
              } else {
                  outBuffer.append(aChar);
              }
          }
      }

      return outBuffer.toString();
  }

    private static char toHex(int nibble) {
      char[] hexDigit   = {
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};
      return hexDigit[(nibble & 0xF)];
  }
    
}
