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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/Constants.java $
 * $Author: Christopher Ho $
 * $Date: 2/17/17 9:58a $
 * $Revision: 25 $
******************************************************************************/



package symbolthree.oracle.excel;

import java.io.File;

public interface Constants {
  public static final String RCS_ID                    =
    "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/Constants.java 25    2/17/17 9:58a Christopher Ho $";
  
    public static String       APPS_RUNAS_MODE           = "APPS_RUNAS_MODE";
    public static String       APPS_RUNAS_PASSWORD       = "APPS_RUNAS_PASSWORD";
    public static String       APPS_RUNAS_RESPONSIBILITY = "APPS_RUNAS_RESPONSIBILITY";
    public static String       APPS_RUNAS_USER           = "APPS_RUNAS_USER";
    public static String       CELL_FORMAT               = "CELL_FORMAT";
    public static String       COLUMN_MAPPING            = "COLUMN_MAPPING";
    public static String       COLUMN_TITLE_FORMAT       = "COLUMN_TITLE_FORMAT";
    public static String       COLUMN_TITLE_ROW          = "COLUMN_TITLE_ROW";
    public static String       COMMIT_AND_EXIT           = "COMMIT_AND_EXIT";
    public static String       CONNECTION_DIRECT         = "DIRECT";
    public static String       CONNECTION_EBS            = "APPLICATIONS";
    public static String       CONNECTION_MODE           = "CONNECTION_MODE";
    public static String       CONTINUE_ON_ERROR         = "CONTINUE_ON_ERROR";
    public static String       CUSTOM_QUERY              = "CUSTOM_QUERY";
    public static String       DATA_WORKSHEET            = "DATA_WORKSHEET";
    public static String       DATE_FORMAT               = "DATE_FORMAT";
    public static String       DATE_MASK                 = "DATE_MASK";
    public static String       DBC_FILE                  = "DBC_FILE";
    public static String       ERROR_HANDLING            = "ERROR_HANDLING";
    public static String       EXZELLENZ_VERSION         = "EXZELLENZ_VERSION";
    public static String       EXZELLENZ_FULL_VERSION    = "EXZELLENZ_FULL_VERSION";
    
    public static String       EXZ_LOG_LEVEL             = "EXZ_LOG_LEVEL";
    public static String       EXZ_LOG_OUTPUT            = "EXZ_LOG_OUTPUT";
    
    public static String       EXZ_LOG_INTERVAL          = "EXZ_LOG_INTERVAL";
    public static String       EXZ_APPLICATION_DIR       = System.getProperty("user.dir");
    public static String       IGONRE_NOT_NULL_COLUMN    = "IGONRE_NOT_NULL_COLUMN";
    
    public static int          LOG_DEBUG                 = 0;
    public static int          LOG_INFO                  = 1;
    public static int          LOG_WARN                  = 2;
    public static int          LOG_ERROR                 = 4;
    public static int          LOG_CONFIRM               = 9;    

    public static String       LOG_DEBUG_COLOR           = "gray";
    public static String       LOG_INFO_COLOR            = "maroon";    
    public static String       LOG_WARN_COLOR            = "orange";    
    public static String       LOG_CONFIRM_COLOR         = "blue";
    public static String       LOG_ERROR_COLOR           = "red";

    public static String       LOG_OUTPUT_FILE           = "FILE";
    public static String       LOG_OUTPUT_FILENAME       = "EXZELLENZ.log";
    public static String       LOG_OUTPUT_SYSTEM_OUT     = "SYSTEM.OUT";
    
    public static String       NLS_LANGUAGE              = "NLS_LANGUAGE";
    public static String       NO_COMMIT_AND_EXIT        = "NO_COMMIT_AND_EXIT";
    public static String       NO_RUNAS_MODE             = "NO_RUNAS_MODE";
    public static String       OPERATION_CLASS_SUFFIX    = "symbolthree.oracle.excel.Operation";
    public static String       OPERATION_DOWNLOAD        = "DOWNLOAD";
    public static String       OPERATION_INSERT          = "INSERT";
    public static String       OPERATION_MODE            = "OPERATION_MODE";
    public static String       OPERATION_TEMPLATE        = "TEMPLATE";
    public static String       OPERATION_UPDATE          = "UPDATE";
    public static String       OPERATION_VALIDATE        = "VALIDATE";
    public static String       ORDER_CLAUSE              = "ORDER_CLAUSE";
    public static String       PARAMETER_WORKSHEET_NAME  = "EXZELLENZ";
    public static String       PASSWORD                  = "PASSWORD";
    public static String       PENDING_DELETE            = "PENDING_DELETE";
    public static String       PENDING_UPDATE            = "PENDING_UPDATE";
    public static String       PLACE_HOLDER              = "$$$$$$$$$$";
    public static String       PROGRAM_LOCALE            = "PROGRAM_LOCALE";
    public static String       RESULT_COLUMN_NAME        = "RESULT_COLUMN_NAME";
    public static String       RESULT_FAILURE            = "RESULT_FAILURE";
    public static String       RESULT_SUCCESS            = "RESULT_SUCCESS";
    public static String       RUN_MODE                  = "exzellenz.mode";
    public static String       JDBC_URL                  = "JDBC_URL";
/*   
 *  As of verison 1.7, users need to provide JDBC URL instead of individual parameters (RAC connection fix) 
    public static String       SERVER                    = "SERVER";
    public static String       PORT                      = "PORT";    
    public static String       SID                       = "SID";
*/    
    public static String       TABLE_NAME                = "TABLE_NAME";
    public static String       USERNAME                  = "USERNAME";
    public static String       USE_RUNAS_MODE            = "USE_RUNAS_MODE";
    public static String       VERSION                   = "VERSION";
    public static String       WHERE_CLAUSE              = "WHERE_CLAUSE";
    public static String       LAST_FILE                 = "LAST_FILE";
    public static String       SAVE_NEW_FILE             = "SAVE_NEW_FILE";
    public static String       RUNMODE_GUI               = "GUI";
    public static String       RUNMODE_CONSOLE           = "CONSOLE";
/*
 *  As of version 1.8, SHOW_ROWCOUNT is a configurable parameter
    public static int          SHOWING_ROWCOUNT          = 100;
 *
**/
    public static String       WINDOW_WIDTH              = "WINDOW_WIDTH";
    public static String       WINDOW_HEIGHT             = "WINDOW_HEIGHT";
    
    public static double       LOWEST_MAJOR_VERSION_ALLOWED  = 1;
    public static double       LOWEST_MINOR_VERSION_ALLOWED  = 11;
    public static int          SXSSF_WINDOW_SIZE         = 1000;
    
    public static String       CONFIRM_OPERATION         = "CONFIRM_OPERATION";  /* new in version 1.10 */
    public static String       CONFIRM_RESPONSE          = "CONFIRM_RESPONSE";   /* new in version 1.10 */
    
    public static String       CREATE_ROWID_COLUMN       = "CREATE_ROWID_COLUMN"; /* new in version 1.11 */
    public static String       NEW_FILE_NAME             = "NEW_FILE_NAME";       /* new in version 1.11 */
    
    public static String       KEEP_PARAMETER_WORKSHEET  = "KEEP_PARAMETER_WORKSHEET"; /* new in version 1.12 */


    
}

