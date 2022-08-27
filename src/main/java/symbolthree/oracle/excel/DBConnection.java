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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/DBConnection.java $
 * $Author: Christopher Ho $
 * $Date: 7/12/16 11:09a $
 * $Revision: 13 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import oracle.apps.fnd.ext.jdbc.datasource.AppsDataSource;
import oracle.jdbc.OracleDriver;

//~--- JDK imports ------------------------------------------------------------


import java.io.File;
import java.sql.*;
import java.util.Properties;

public class DBConnection implements Constants {
    private static DBConnection instance = null;
    private Connection          connection;

    protected DBConnection(String urlWithCredentials, Properties sessionProp, boolean useRunAsMode)
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

        if (useRunAsMode) {
            connectRunAsUser();
        }
    }

    protected DBConnection(String url, String username, String password, boolean useRunAsMode) throws EXZException {
        try {
            DriverManager.registerDriver(new OracleDriver());
            connection = DriverManager.getConnection(url, username, password);
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

        if (useRunAsMode) {
            connectRunAsUser();
        }
    }

    protected DBConnection(String appUser, String appPassword, File dbcFile, Properties sessionProp,
                           boolean useRunAsMode)
            throws EXZException {
        try {
            AppsDataSource ads = new AppsDataSource();

            ads.setDescription("EXZELLENZ");
            ads.setUser(appUser);
            ads.setPassword(appPassword);
            ads.setDbcFile(dbcFile.getAbsolutePath());

            connection = ads.getConnection();
            if (sessionProp != null) {
                ads.setConnectionProperties(sessionProp);
            }

            connection.setAutoCommit(false);

            Statement stmt     = connection.createStatement();
            String    language = EXZParams.instance().getValue(NLS_LANGUAGE);

            stmt.execute("alter session set NLS_LANGUAGE='" + language + "'");
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        }

        if (useRunAsMode) {
            connectRunAsUser();
        }
    }

    public static DBConnection getInstance(String url, String username, String password, boolean useRunAsMode)
            throws EXZException {
        
        if (EXZHelper.isEmpty(url) ||
            EXZHelper.isEmpty(username) ||
            EXZHelper.isEmpty(password)) {
          throw new EXZException(EXZI18N.inst().get("ERR.DATABASE_PARAMS"));
        }
      
        if (instance == null) {
            instance = new DBConnection(url, username, password, useRunAsMode);
        }

        return instance;
    }

    public static DBConnection getInstance(String urlWithCredential, Properties sessionProp, boolean useRunAsMode)
            throws EXZException {
      
        if (instance == null) {
            instance = new DBConnection(urlWithCredential, sessionProp, useRunAsMode);
        }

        return instance;
    }

    public static DBConnection getInstance(String appUser, String appPassword, File dbcFile, Properties sessionProp,
            boolean useRunAsMode)
            throws EXZException {
      
        if (EXZHelper.isEmpty(appUser) ||
            EXZHelper.isEmpty(appPassword)) {
          throw new EXZException(EXZI18N.inst().get("ERR.DATABASE_PARAMS"));
        }
      
        if (instance == null) {
            instance = new DBConnection(appUser, appPassword, dbcFile, sessionProp, useRunAsMode);
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
                                           
    
    private void connectRunAsUser() throws EXZException {
        String appsRunAsUser = EXZParams.instance().getValue(APPS_RUNAS_USER);
        String appsRunAsPwd  = EXZParams.instance().getValue(APPS_RUNAS_PASSWORD);
        String appsResp      = EXZParams.instance().getValue(APPS_RUNAS_RESPONSIBILITY);
        String language      = EXZParams.instance().getValue(NLS_LANGUAGE);

        if (EXZHelper.isEmpty(appsRunAsUser) || EXZHelper.isEmpty(appsRunAsPwd) || EXZHelper.isEmpty(appsResp)
                || EXZHelper.isEmpty(language)) {
            throw new EXZException("All RunAs paramaeters are required.");
        }

        checkRunAsCredentials();

        String sql =
            "select a.user_id || ',' || b.responsibility_id || ',' || c.application_id"
          + "     , f.application_short_name "		
          + "  from fnd_user a"
          + "     , fnd_user_resp_groups b"
          + "     , fnd_responsibility c"
          + "     , fnd_responsibility_tl d"
          + "     , fnd_languages e "
          + "     , fnd_application f"
          + " where c.responsibility_id=d.responsibility_id "
          + "   and B.RESPONSIBILITY_ID=d.responsibility_id "
          + "   and B.user_id=to_char(a.user_id) "
          + "   and upper(a.user_name)= ? "
          + "   and upper(d.responsibility_name) = ? "
          + "   and d.language=e.language_code "
          + "   and e.nls_language=?"
          + "   and f.application_id=c.responsibility_application_id";

        try {
            PreparedStatement prepStmt = connection.prepareStatement(sql);

            prepStmt.setString(1, appsRunAsUser.toUpperCase());
            prepStmt.setString(2, appsResp.toUpperCase());
            prepStmt.setString(3, language);

            ResultSet rs         = prepStmt.executeQuery();
            String    sessionCtx   = null;
            String    appShortName = null; 

            while (rs.next()) {
            	sessionCtx   = rs.getString(1);
            	appShortName = rs.getString(2);
            }
            rs.close();

            Statement stmt = connection.createStatement();

            if (sessionCtx != null) {
                sql = "BEGIN fnd_global.apps_initialize(" + sessionCtx + "); END;";
                stmt.execute(sql);
            }
            stmt.close();
            
            sql = "SELECT RELEASE_NAME FROM FND_PRODUCT_GROUPS";
            rs = connection.createStatement().executeQuery(sql);
            rs.next();
            String releaseName = rs.getString(1);
            rs.close();
            if (releaseName.startsWith("12.")) {
              sql = "BEGIN MO_GLOBAL.INIT('" + appShortName + "'); EXCEPTION WHEN OTHERS THEN NULL; END;";
              stmt.execute(sql);
            }
            stmt.close();
            
        } catch (SQLException sqle) {
            throw new EXZException(EXZI18N.inst().get("ERR.RUNAS_CREDENTIALS"));
        }
    }

    private void checkRunAsCredentials() throws EXZException {
        String appsRunAsUser     = EXZParams.instance().getValue(APPS_RUNAS_USER);
        String appsRunAsPwd      = EXZParams.instance().getValue(APPS_RUNAS_PASSWORD);
        String sql               = "select fnd_web_sec.validate_login(?,?) from dual";
        String isValidCredential = "N";

        try {
            PreparedStatement prepStmt = connection.prepareStatement(sql);

            prepStmt.setString(1, appsRunAsUser);
            prepStmt.setString(2, appsRunAsPwd);

            ResultSet rs = prepStmt.executeQuery();

            rs.next();
            isValidCredential = rs.getString(1);
            rs.close();
        } catch (SQLException sqle) {
            throw new EXZException(sqle);
        }

        if (!isValidCredential.equals("Y")) {
            throw new EXZException(EXZI18N.inst().get("ERR.RUNAS_CREDENTIALS"));
        } else {
            return;
        }
    }
    
    public void clear() {
    	instance = null;
    }
}
