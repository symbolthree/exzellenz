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

import oracle.jdbc.OracleDriver;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

//~--- JDK imports ------------------------------------------------------------


import java.util.Properties;

public class DBConnection implements Constants {
    private static DBConnection instance = null;
    private Connection          connection;

    protected DBConnection(String urlWithCredentials, Properties sessionProp)
            throws EXZException {
        try {
            DriverManager.registerDriver(new OracleDriver());
            connection = DriverManager.getConnection(urlWithCredentials, sessionProp);
            connection.setAutoCommit(false);

            Statement stmt     = connection.createStatement();
            String    language = EXZParams.instance().getValue(NLS_LANGUAGE);

            stmt.execute("alter session set NLS_LANGUAGE='" + language + "'");

            String sql = "BEGIN DBMS_APPLICATION_INFO.set_module(" + "module_name => '"
                         + sessionProp.getProperty("v$session.module") + "'," + "action_name => '"
                         + sessionProp.getProperty("v$session.action") + "');END;";

            stmt.execute(sql);
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        }
    }

    protected DBConnection(String url, String username, String password) throws EXZException {
        try {
            DriverManager.registerDriver(new OracleDriver());
            connection = DriverManager.getConnection(url, username, getPassword(password));
            connection.setAutoCommit(false);

            Statement stmt     = connection.createStatement();
            String    language = EXZParams.instance().getValue(NLS_LANGUAGE);
            
            if (language==null || language.equals("")) {
            	language = "AMERICAN";
            }

            stmt.execute("alter session set NLS_LANGUAGE='" + language + "'");
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        }
    }

    public static DBConnection getInstance(String url, String username, String password)
            throws EXZException {
        
        if (EXZHelper.isEmpty(url) ||
            EXZHelper.isEmpty(username) ||
            EXZHelper.isEmpty(password)) {
          throw new EXZException(EXZI18N.inst().get("ERR.DATABASE_PARAMS"));
        }
      
        if (instance == null) {
            instance = new DBConnection(url, username, getPassword(password));
        }

        return instance;
    }

    public static DBConnection getInstance(String urlWithCredential, Properties sessionProp)
            throws EXZException {
      
        if (instance == null) {
            instance = new DBConnection(urlWithCredential, sessionProp);
        }

        return instance;
    }

    public static DBConnection getInstance() throws SQLException {
        if (instance == null) {
            throw new SQLException("DBConnection is not instantiated.");
        } else {
            return instance;
        }
    }

    
    public Connection getConnection() {
        return connection;
    }

    
    public void releaseConnection() {
      if (connection != null) {
        try {
          connection.close();
        } catch (SQLException sqle) {
          // do nothing
        }
        connection = null;
      }
    }
    
    private static String getPassword(String pwd) {
        if ((pwd != null) && (pwd.length() == ENCRYPTED_PASSWORD_LENGTH)) {
            return Security.getInstance().decryptPwd(pwd);
        } else {
            return pwd;
        }    	
    }

}
