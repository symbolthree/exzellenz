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

import java.io.*;
import java.sql.SQLException;
import java.util.*;

public class EXZELLENZ implements Runnable, Constants {
    private Hashtable<String, Cell> specialCells = new Hashtable<String, Cell>();
    private File                    excelFile;
    private FileInputStream         fis;
    private String                  excelFormat;    // XLS, XLSX or XLSB
    private String                  fileFromGUI;            
    
    public static void main(String[] args) {
        System.setProperty(RUN_MODE, RUNMODE_CONSOLE);
        EXZHelper.getVersion();

        if (args.length != 1) {
            printUsage();
            System.exit(1);
        } else {
            printHeader();
        }

        EXZELLENZ exzellenz = new EXZELLENZ();

        try {
        	exzellenz.start(args);
        } catch (Exception e) {
            EXZHelper.logError(e);
        }
    }

    private void start(String[] args) throws EXZException, IOException {
        EXZHelper.initializeLogging();
        start(args[0]);
    }

    // this method is called from GUI
    public void start(String _file) throws EXZException {
        if (!EXZELLENZ.checkJRE()) {
            EXZHelper.log(LOG_ERROR, "Please use JRE 1.8 or higher");

            throw new EXZException();
        }

        try {
            excelFile = new File(_file);

            if (!excelFile.exists()) {
                throw new EXZException("Cannot find file " + _file);
            }

            // check Excel 97 or 2007 format
            String ext = _file.substring(_file.lastIndexOf(".") + 1, _file.length());

            if (ext.equalsIgnoreCase("xls")) {
                excelFormat = "XLS";
            } else if (ext.equalsIgnoreCase("xlsx")) {
                excelFormat = "XLSX";
            } else if (ext.equalsIgnoreCase("xlsb")) {
                excelFormat = "XLSB";
            } else {
                throw new EXZException("Invalid file extension (" + ext + ")");
            }

            fis = new FileInputStream(excelFile);
            Workbook        wb  = null;

            if (excelFormat.equals("XLS")) {
                wb = new HSSFWorkbook(fis);
            } else if (excelFormat.equals("XLSX") || excelFormat.equals("XLSB")) {
                wb = new XSSFWorkbook(fis);
            }

            int noOfSheet = wb.getNumberOfSheets();

            EXZHelper.log(LOG_DEBUG, "noOfSheet: " + noOfSheet);

            int idx = wb.getSheetIndex(PARAMETER_WORKSHEET_NAME);

            if (idx < 0) {
                wb.close();
                throw new EXZException(EXZI18N.inst().get("ERR.SHEET_NOTFOUND", PARAMETER_WORKSHEET_NAME));
            } else {
                EXZHelper.log(LOG_DEBUG, PARAMETER_WORKSHEET_NAME + " parameter worksheet present");
            }

            Sheet exzSheet = wb.getSheet(PARAMETER_WORKSHEET_NAME);
            int   colNo    = 2;
            int   lastRow  = exzSheet.getLastRowNum();

            EXZHelper.log(LOG_DEBUG, "Last Row = " + (lastRow + 1));

            // clear all parameters first
            EXZParams.instance().clear();
            
            Hashtable<String, String> customColumnMap = new Hashtable<String, String>();

            boolean customMappingFound = false;
            
            String paramType = null;
            String param     = null;
            String value     = null;
            
            while (colNo <= lastRow + 1) {
                paramType = EXZHelper.readString(wb, exzSheet, colNo, 1);
                param     = EXZHelper.readString(wb, exzSheet, colNo, 2);
                value     = EXZHelper.readString(wb, exzSheet, colNo, 3);

                if (EXZHelper.isEmpty(paramType) && EXZHelper.isEmpty(param) && EXZHelper.isEmpty(value)) {
                    break;
                }

                if (!EXZHelper.isEmpty(paramType) && 
                	!paramType.equals(COLUMN_MAPPING) &&
                	!paramType.equals(PLACE_HOLDER)) {
                	
                    if (param.endsWith(PASSWORD)) {
                        EXZHelper.log(LOG_DEBUG, colNo + "-" + param + ":" + EXZHelper.masking(value, "*"));
                    } else {
                        EXZHelper.log(LOG_DEBUG, colNo + "-" + param + ":" + value);
                    }

                    EXZParams.instance().setValue(param, value);
                }

                if (!EXZHelper.isEmpty(paramType) && 
                	paramType.equals(COLUMN_MAPPING) &&
                	!EXZHelper.isEmpty(value)) {
                	
                	if (!customMappingFound) {
                		customMappingFound = true;
                	} else {
                      EXZHelper.log(LOG_DEBUG, "Column Mapping - " + param + ":" + value);
                      customColumnMap.put(param, value);
                	}
                }

                if (!EXZHelper.isEmpty(paramType) && paramType.equals(CELL_FORMAT)) {
                    Cell cell = exzSheet.getRow(colNo - 1).getCell(2);

                    specialCells.put(param, cell);
                }

                colNo++;
            }

            // check Excel template and program version compatibility
            String templateVer = EXZParams.instance().getValue(VERSION);
            int templateMajorVer = Integer.parseInt(templateVer.split("\\.")[0]);
            int templateMinorVer = Integer.parseInt(templateVer.split("\\.")[1]);
            
            if (templateMajorVer < LOWEST_MAJOR_VERSION_ALLOWED) {
                EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.VERSION", LOWEST_MAJOR_VERSION_ALLOWED + "." + LOWEST_MINOR_VERSION_ALLOWED));
                throw new EXZException();
            }
            
            if (templateMajorVer == LOWEST_MAJOR_VERSION_ALLOWED && templateMinorVer < LOWEST_MINOR_VERSION_ALLOWED) {            
                EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.VERSION", LOWEST_MAJOR_VERSION_ALLOWED + "." + LOWEST_MINOR_VERSION_ALLOWED));
                throw new EXZException();
            }

            // In here, the OPERATION_MODE is found. We can use the factory to create different operations.
            String opMode = EXZParams.instance().getValue(OPERATION_MODE);

            EXZHelper.log(LOG_INFO, "Operation : " + opMode);
            
            String tableName  = EXZParams.instance().getValue(TABLE_NAME);
            String owner      = EXZParams.instance().getValue(OWNER);
            String username   = EXZParams.instance().getValue(USERNAME);
            
            if (owner != null && ! owner.equals("")) {
            	EXZHelper.log(LOG_INFO, "Table : " + owner + "." + tableName);
            } else {
            	EXZHelper.log(LOG_INFO, "Table : " + username + "." + tableName);
            }
            
            // if custom query is used, worksheet name must be present
            if (EXZParams.instance().getValue(CUSTOM_QUERY) != null && 
            	EXZParams.instance().getValue(DATA_WORKSHEET)==null) {
                EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.CUSTOM_QUERY_SHEET"));
                throw new EXZException();
            }
            
            // CUSTOM_QUERY and TABLE_NAME is mutually exclusive
            if (EXZParams.instance().getValue(CUSTOM_QUERY) != null && 
            	EXZParams.instance().getValue(TABLE_NAME)   != null) {
                EXZHelper.log(LOG_ERROR, EXZI18N.inst().get("ERR.CUSTOM_QUERY_CONFLICT"));
                throw new EXZException();
            }            

            // prompt for confirmation
            if (EXZParams.instance().getValue(CONFIRM_OPERATION) != null &&
           		EXZParams.instance().getValue(CONFIRM_OPERATION).equals("Y")) {
            	
            	if (System.getProperty(RUN_MODE).equals(RUNMODE_GUI)) {
                  EXZHelper.log(LOG_CONFIRM, EXZI18N.inst().get("MSG.CONFIRM_CONTINUE_GUI"));            		
            	} else {
              	  EXZHelper.log(LOG_CONFIRM, EXZI18N.inst().get("MSG.CONFIRM_CONTINUE_CONSOLE"));            		
            	}
            	
                if (!showConfirmation()) {
                	EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.CONFIRM_CANCELLED"));
                	return;
                };
            }
            
            OperationFactory factory   = new OperationFactory(opMode);
            Operation        operation = factory.getOperation();

            if (operation.checkDBConnection()) {
                operation.setWorkbook(wb);
                operation.setSpecialCells(specialCells);
                operation.defineTableName();
                operation.createColumnMapping(customColumnMap);
                operation.setExcelFile(excelFile);
                operation.doOperation();
                operation.postOperation();

                if (operation.isSaveFileRequired()) {
                    EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.PREPARE_WRITE_FILE"));                  
                    fis.close();
                    operation.writeWorkbook();
                }
                
                DBConnection.getInstance().releaseConnection();
                DBConnection.getInstance().clear();
                EXZHelper.log(LOG_INFO, EXZI18N.inst().get("MSG.PROCESS_DONE"));
            }
        } catch (Exception e) {
            try {
              if (fis != null) fis.close();
              DBConnection.getInstance().releaseConnection();
              DBConnection.getInstance().clear();
            } catch (IOException ioe) {
              // do nothing
            } catch (SQLException ioe) {
              // do nothing
            } 
            throw new EXZException(e);
        }
    }

    private static boolean checkJRE() {
        double ver = Double.parseDouble(System.getProperty("java.specification.version"));

        if (ver < 1.7) {
            return false;
        } else {
            return true;
        }
    }

    private static void printHeader() {
    	System.out.println("~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~");    	
    	try {
	    	InputStream banner = EXZELLENZ.class.getResourceAsStream("/symbolthree/oracle/excel/banner.txt");
	    	BufferedReader br = new BufferedReader(new InputStreamReader(banner));
	    	String line;
	        while ((line = br.readLine()) != null) {
	            System.out.println(line);
	        }
    	} catch (IOException ioe) {
    		EXZHelper.log(LOG_INFO, "unable to load banner");
    	}
        System.out.println("EXZELLENZ " + EXZHelper.getVersionWithTimestamp());
        System.out.println(EXZHelper.getAuthorLine());
        System.out.println("~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~^~");
    }

    private static void printUsage() {
    	printHeader();
        System.out.println("EXZELLENZ.bat/sh [Excel file name with full path]");
        System.out.println("e.g. EXZCELLENZ.bat \"C:\\My Document\\example.xls\"");
        System.out.println("Please use double quote if the path or file name contains space");
    }
    
    // This method is invoked from GUI
    public void run() {
      try {
        this.start(fileFromGUI);
      } catch (EXZException e) {
        EXZHelper.logError(e);
      }
    }

    // This method is invoked from GUI
    protected void setFile(String _file) {
      fileFromGUI = _file;
    }

    // show confirmation dialog
    private boolean showConfirmation() {
      if (System.getProperty(EXZ_LOG_OUTPUT).equals(LOG_OUTPUT_SYSTEM_OUT)) {
	    if (System.getProperty(RUN_MODE).equals(RUNMODE_GUI)) {
  		  return showConfirmationGUI();
	    } else {
		 return showConfirmationConsole();    		
	    }
      } else{
    	  return true;
      }
    }
    
    private boolean showConfirmationGUI() {
	  boolean keepWaiting = true;
	  boolean rtnVal = false;
	  
	  while (keepWaiting) {
	    if (System.getProperty(CONFIRM_RESPONSE) == null || 
    		System.getProperty(CONFIRM_RESPONSE).equals("")) {
  		  try {
			Thread.sleep(500);
		  } catch (InterruptedException e) {
			e.printStackTrace();
			return false;
		  }    			  
	    } else {
		  keepWaiting = false;
	      if (System.getProperty(CONFIRM_RESPONSE).equals("Y")) {
	    	rtnVal = true;
	    	break;
	      } else {
    	    rtnVal = false;
	    	break;
	      }    
	    } // response decided
	  } // response provided
	  System.clearProperty(CONFIRM_RESPONSE);
	  return rtnVal;
	}
    
    private boolean showConfirmationConsole() {
	  Scanner scanner = new Scanner(System.in);
  	  String rtnVal = scanner.nextLine();
	  scanner.close();    		
	  if (rtnVal.equalsIgnoreCase("Y")) {
		return true;
	  } else {
		return false;
	  }
	}
}
