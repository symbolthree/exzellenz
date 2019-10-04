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
 * $Archive: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZHelper.java $
 * $Author: Christopher Ho $
 * $Date: 2/15/17 6:14a $
 * $Revision: 24 $
******************************************************************************/

package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.*;
import java.util.*;

public class EXZHelper implements Constants {
    public static final String RCS_ID =
        "$Header: /TOOL/EXZELLENZ/src/symbolthree/oracle/excel/EXZHelper.java 24    2/15/17 6:14a Christopher Ho $";
    final static private SimpleDateFormat timeFormat = new SimpleDateFormat("yyMMdd.HHmmss");
    public static String                  MAJOR_VER;
    public static String                  MINOR_VER;
    private static FileWriter             logWriter;

    public static String readString(Workbook wb, Sheet sheet, int rowNo, int colNo) {

        // If the cell value is a numeric number, it is converted to integer
        // The return value is "" (empty string) if it is null
        // If you need to read a double value, please use readDoubleFromSheet
        String rtnValue = "";

        try {
            if (wb instanceof HSSFWorkbook) {
                HSSFRow row = ((HSSFSheet) sheet).getRow(rowNo - 1);

                if (row != null) {
                    HSSFCell cell = row.getCell(colNo - 1);

                    if (cell != null) {
                        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                            rtnValue = Integer.toString((int) cell.getNumericCellValue());
                            
                        } else if (cell.getCellTypeEnum() == CellType.STRING) {
                            HSSFRichTextString richText = new HSSFRichTextString();

                            richText = cell.getRichStringCellValue();
                            rtnValue = richText.getString().trim();
                        
                        } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                        	rtnValue = readFormulaAsString(wb, cell);
                        }
                    }
                }
            } else if (wb instanceof XSSFWorkbook) {
                XSSFRow row = ((XSSFSheet) sheet).getRow(rowNo - 1);

                if (row != null) {
                    XSSFCell cell = row.getCell(colNo - 1);

                    if (cell != null) {
                        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                            rtnValue = Integer.toString((int) cell.getNumericCellValue());
                            
                        } else if (cell.getCellTypeEnum() == CellType.STRING) {
                            XSSFRichTextString richText = new XSSFRichTextString();

                            richText = cell.getRichStringCellValue();
                            rtnValue = richText.getString().trim();
                        
                        } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                        	rtnValue = readFormulaAsString(wb, cell);
                        }
                        
                    }
                }
            }
        } catch (Exception e) {
            EXZHelper.log(LOG_ERROR, "Error in reading cell " + EXZHelper.number2Letter(colNo) + ":" + rowNo);
            rtnValue = null;
        }

        return rtnValue;
    }

    public static double readDouble(Workbook wb, Sheet sheet, int rowNo, int colNo) throws EXZException {
        double rtnValue = Double.MIN_VALUE;
        Row    row      = sheet.getRow(rowNo - 1);

        if (row != null) {
            Cell cell = row.getCell(colNo - 1);

            if (cell != null) {
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    rtnValue = cell.getNumericCellValue();
                    
                } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                    rtnValue = readFormulaAsDouble(wb, cell);
                    
                } else if (cell.getCellTypeEnum() == CellType.BLANK) {

                    // do nothing.  It will return min. value which is indicated as blank
                } else if (cell.getCellTypeEnum() == CellType.STRING) {
                	try {
                        rtnValue = Double.parseDouble(cell.getRichStringCellValue().getString());
                    } catch (NumberFormatException nfe) {
                        throw new EXZException("Cell " + EXZHelper.number2Letter(colNo) + ":" + rowNo
                                               + " is not a number");
                    }
                }
            }
        }

        return rtnValue;
    }


    private static String readFormulaAsString(Workbook wb, Cell cell) {
        FormulaEvaluator evaluator = null;

        if (wb instanceof HSSFWorkbook) {
            evaluator = new HSSFFormulaEvaluator((HSSFWorkbook) wb);
        } else if (wb instanceof XSSFWorkbook) {
            evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
        }

        return evaluator.evaluate(cell).getStringValue();
    }    
    
    private static double readFormulaAsDouble(Workbook wb, Cell cell) {
        FormulaEvaluator evaluator = null;

        if (wb instanceof HSSFWorkbook) {
            evaluator = new HSSFFormulaEvaluator((HSSFWorkbook) wb);
        } else if (wb instanceof XSSFWorkbook) {
            evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
        }

        return evaluator.evaluate(cell).getNumberValue();
    }

    public static void setCellStyle(HSSFWorkbook wb, HSSFSheet sheet, int rowNo, int colNo, short bgColor,
                                    short fontColor, String cellText) {
        HSSFRow row = sheet.getRow(rowNo - 1);

        if (row != null) {
            HSSFCell      cell      = row.createCell(colNo - 1);
            HSSFCellStyle cellStyle = wb.createCellStyle();

            cellStyle.setFillBackgroundColor(bgColor);
            cellStyle.setFillBackgroundColor(fontColor);

            HSSFRichTextString richText = new HSSFRichTextString(cellText);

            cell.setCellValue(richText);
            cell.setCellStyle(cellStyle);
        }
    }

    public static void successCell(HSSFWorkbook workbook, HSSFSheet sheet, int rowNo, int colNo) {
        setCellStyle(workbook, sheet, rowNo, colNo, HSSFColor.BRIGHT_GREEN.index, HSSFColor.BLACK.index, "Uploaded");
    }

    public static boolean isEmptyRow(HSSFSheet sheet, int rowNo, int colNo) {
        return false;
    }

    public static java.sql.Date readDate(Workbook wb, Sheet sheet, int rowNo, int colNo) throws Exception {
        try {
            Row row = sheet.getRow(rowNo - 1);

            if (row != null) {
                Cell cell = row.getCell(colNo - 1);

                if (cell != null) {
                    EXZHelper.log(LOG_DEBUG, "Cell type:" + cell.getCellTypeEnum());
                    
                    if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        // best situation: cell is date formatted
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            java.util.Date jDate = cell.getDateCellValue();

                            return new java.sql.Date(jDate.getTime());
                        }
                    
                    } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                    	FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                    	CellValue cv = evaluator.evaluate(cell);
                    	java.util.Date jDate = HSSFDateUtil.getJavaDate(cv.getNumberValue());
                    	return new java.sql.Date(jDate.getTime());
                        
                    } else if (cell.getCellTypeEnum() == CellType.STRING) {
                        String strValue = cell.getRichStringCellValue().getString();
                        /* TODO SYSDATE casting */
                        
                        if (strValue.toUpperCase().equals("SYSDATE")) {
                        	return new java.sql.Date(Calendar.getInstance().getTime().getTime());
                        }
                        
                        // if the cell is stored as string, using date mask value to cast to back to calendar date
                        String dataMask = EXZParams.instance().getValue(DATE_MASK);

                        if (!EXZHelper.isEmpty(dataMask)) {
                            SimpleDateFormat cellDateFormat = new SimpleDateFormat(dataMask);
                            java.util.Date   jDate          = cellDateFormat.parse(strValue);

                            return new java.sql.Date(jDate.getTime());
                        } else {

                            // no masking provided
                            EXZHelper.log(LOG_ERROR, "Please specify the DATE_MASK value.");

                            throw new EXZException("Cell value is string but no date mask provided.");
                        }
                    
                    } else if (cell.getCellTypeEnum() == CellType.BLANK) {
                        return null;
                    
                    } else {
                        EXZHelper.log(LOG_ERROR, "Please apply date format to the cell.");

                        throw new EXZException("Date expected but format is not correct.");
                    }
                }
            }
        } catch (Exception e) {
            throw new EXZException(e);
        }

        return null;
    }

    public static int letter2Number(String colRef) {
        String allLetter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        int    letterLen = colRef.length();
        int    output    = 0;
        String letters   = reverse(colRef);

        for (int i = 0; i < letterLen; i++) {
            output = output + (allLetter.indexOf(letters.charAt(i)) + 1) * (int) Math.pow(26.0d, (double) i);
        }

        return output;
    }

    public static String number2Letter(int colNumber) {
        String letters = "";

        // the largest possible value = 2^14 = 16384 (XFD)
        int    col       = colNumber;
        String allLetter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        int    letterLen = (int) Math.ceil((Math.log((double) col) / Math.log(26.0d)));

        for (int i = 0; i < letterLen; i++) {
            int remainder = col % 26;

            letters += allLetter.charAt((int) remainder - 1);
            col     = (col - remainder) / 26;
        }

        if (letters.equals("")) {
            return "A";
        } else {
            return reverse(letters);
        }
    }

    private static String reverse(String str) {
        String output    = "";
        int    letterLen = str.length();

        for (int i = letterLen - 1; i > -1; i--) {
            output += str.charAt(i);
        }

        return output;
    }

    public static boolean isEmpty(String str) {
        if ((str == null) || (str.trim().length() == 0)) {
            return true;
        } else {
            return false;
        }
    }

    public static String getStackTrace(Throwable t) {
        StringWriter sw = new StringWriter();
        PrintWriter  pw = new PrintWriter(sw, true);

        t.printStackTrace(pw);
        pw.flush();
        sw.flush();

        return sw.toString();
    }

    public static void logError(Throwable t) {
        if (System.getProperty(EXZ_LOG_LEVEL).equals(logLevelStr(LOG_DEBUG))) {
            log(LOG_ERROR, getStackTrace(t));
        } else {
            log(LOG_ERROR, t.getLocalizedMessage());
        }
    }

    public static void log(int logLevel, String logMsg) {
        int    logThreshold = logLevelInt(System.getProperty(EXZ_LOG_LEVEL));
        String logOutput    = System.getProperty(EXZ_LOG_OUTPUT);
        String logLine      = timeFormat.format(new java.util.Date()) + 
                              " [" + logLevelStr(logLevel) + "] - " + logMsg;

        if ((logLevel >= logThreshold) || (logLevel == LOG_ERROR)) {
            
            if (logOutput.equals(LOG_OUTPUT_SYSTEM_OUT))  {
            	
                if (System.getProperty(RUN_MODE).equals(RUNMODE_GUI)) {
                	if (logLevel==LOG_DEBUG)   logMsg = "<font color='" + LOG_DEBUG_COLOR   + "'>" +  logMsg + "</font>";                	
                	if (logLevel==LOG_INFO)    logMsg = "<font color='" + LOG_INFO_COLOR    + "'>" +  logMsg + "</font>";
                	if (logLevel==LOG_WARN)    logMsg = "<font color='" + LOG_WARN_COLOR    + "'>" +  logMsg + "</font>";
                	if (logLevel==LOG_ERROR)   logMsg = "<font color='" + LOG_ERROR_COLOR   + "'>" +  logMsg + "</font>";
                	if (logLevel==LOG_CONFIRM) logMsg = "$$$" +  logMsg;                	
                	
                	System.out.println(logMsg);                		
                	
                } else {
                	
                	if (logLevel==LOG_CONFIRM) {
                		System.out.print(logLine);
                	} else {
                		System.out.println(logLine);                		
                	}

                }
            }
            
            if (logOutput.equals(LOG_OUTPUT_FILE)) {
                try {
                    //logWriter.write(logLine + System.lineSeparator());
                	logWriter.write(logLine + System.getProperty("line.separator"));
                    logWriter.flush();
                } catch (IOException ioe) {
                    ioe.printStackTrace();
                }
            }
        }
    }

    public static String logLevelStr(int logLevel) {
        if (logLevel == LOG_DEBUG) {
            return "DEBUG";
        }

        if (logLevel == LOG_INFO) {
            return "INFO";
        }

        if (logLevel == LOG_WARN) {
            return "WARN";
        }

        if (logLevel == LOG_ERROR) {
            return "ERROR";
        }

        if (logLevel == LOG_CONFIRM) {
            return "CONFIRM";
        }
        
        return String.valueOf("UNKNOWN");
    }

    public static int logLevelInt(String logLevel) {
        if (logLevel.equals("DEBUG")) {
            return LOG_DEBUG;
        }

        if (logLevel.equals("INFO")) {
            return LOG_INFO;
        }

        if (logLevel.equals("WARN")) {
            return LOG_WARN;
        }

        if (logLevel.equals("ERROR")) {
            return LOG_ERROR;
        }

        return LOG_ERROR;
    }

    public static void initializeLogging() throws IOException {

        // default logging is: log level = INFO; log output = SYSTEM.OUT
        String val = null;

        val = EXZProp.instance().getStr(EXZ_LOG_LEVEL);

        if (EXZHelper.isEmpty(val)) {
            System.setProperty(EXZ_LOG_LEVEL, "INFO");
        } else {
            System.setProperty(EXZ_LOG_LEVEL, val);
        }

        val = EXZProp.instance().getStr(EXZ_LOG_OUTPUT);

        if (EXZHelper.isEmpty(val)) {
            System.setProperty(EXZ_LOG_OUTPUT, LOG_OUTPUT_SYSTEM_OUT);
        } else {
            System.setProperty(EXZ_LOG_OUTPUT, val);
        }

        if (val.equals(LOG_OUTPUT_FILE)) {
            //File logFile = new File(System.getProperty("user.dir"), LOG_OUTPUT_FILENAME);
            File logFile = new File(EXZ_APPLICATION_DIR, LOG_OUTPUT_FILENAME);

            if (logFile.exists()) {
                logFile.delete();
            }

            logWriter = new FileWriter(logFile);
        }
    }

    public static String masking(String str, String mask) {
        StringBuffer sb = new StringBuffer();

        for (int i = 0; i < str.length(); i++) {
            sb.append(mask);
        }

        return sb.toString();
    }

    // loaded the program version and build no. as a system variable EXZELLENZ_VERSION
    // 0 = all; 1=major; 2=minor
    public static void getVersion() {
        InputStream is        = EXZHelper.class.getResourceAsStream("/build.properties");
        Properties  buildProp = new Properties();

        try {
            buildProp.load(is);
        } catch (Exception e) {

            // do nothing
        }

        MAJOR_VER = buildProp.getProperty("build.version");
        MINOR_VER = buildProp.getProperty("build.number");

        String ver = MAJOR_VER + " build " + MINOR_VER;

        System.setProperty(EXZELLENZ_VERSION, MAJOR_VER);
        System.setProperty(EXZELLENZ_FULL_VERSION, ver);
    }

	public static String getVersionWithTimestamp() {
        InputStream is        = EXZHelper.class.getResourceAsStream("/build.properties");
        Properties  buildProp = new Properties();
        try {
            buildProp.load(is);
        } catch (Exception e) {
            // do nothing
        }
        String MAJOR_VER = buildProp.getProperty("build.version");
        String MINOR_VER = buildProp.getProperty("build.number");
        String TIMSETAMP = buildProp.getProperty("build.time").substring(0,10);
        String ver = "Version " + MAJOR_VER + " build " + MINOR_VER + " (" + TIMSETAMP + ")";
        return ver;
	}    
    
	public static String getAuthorLine() {
		return "Copyright(c) 2010-19 Christopher.Ho@symbolthree.com";
	}
	
    public static String getExtension(File f) {
      try {
          if (f != null) {
              String ext = null;
              String s   = f.getName();
              int    i   = s.lastIndexOf('.');

              if ((i >= 0) && (i < s.length() - 1)) {
                  ext = s.substring(i + 1).toLowerCase();
              }

              return ext;
          } else {
              return null;
          }
      } catch (Exception e) {
          return null;
      }
  }       
    
    
    public static File getNewFile(File _file) {
        String newFileName = EXZParams.instance().getValue(NEW_FILE_NAME);
        if (newFileName != null && ! newFileName.trim().equals("")) {
          return getNewFile2(_file);
        } else {
          return getNewFile1(_file);
        }
    }
    
    public static File getNewFile2(File _file) {
        String _fileName = EXZParams.instance().getValue(NEW_FILE_NAME);
        String dir = _file.getParent();
        String ext = getExtension(_file);
        
        _fileName = _fileName.trim();
        
    	Date currDate = new Date();
      	String result = "";
      	
        EXZHelper.log(LOG_DEBUG, "expected file name = " + _fileName);
      	
      	int pos1 = _fileName.indexOf("$D{");
      	int pos2 = 0;
      	  
      	if (pos1 >= 0) {
      	  boolean endFound = false;
      	  int i = 0;
      	  while (! endFound) {
      		if (_fileName.substring(pos1+i, pos1+i+1).equals("}")) {
      			pos2 = pos1+i;
      			endFound = true;
      		} else {
      			i++;
      		}
      	}  
      	
      	String timestamp = _fileName.substring(pos1+3, pos2);
      	EXZHelper.log(LOG_DEBUG, "timestamp = " + timestamp);
      	result = timestamp;
      	try {
          SimpleDateFormat format = new SimpleDateFormat(timestamp);
          result = format.format(currDate);
          EXZHelper.log(LOG_DEBUG, "timestamp resolved = " + result);          
      	} catch (Exception e) {
        }
     }
     
     if (pos1 < 0) pos1 = 0;
     if (pos2 == 0) pos2 = -1;
     
     EXZHelper.log(LOG_DEBUG, "part 1 = " + _fileName.substring(0, pos1));
     EXZHelper.log(LOG_DEBUG, "part 2 = " + result);
     EXZHelper.log(LOG_DEBUG, "part 3 = " + _fileName.substring(pos2+1));
     
     String fileName = _fileName.substring(0, pos1)  +  result + _fileName.substring(pos2+1);
     File file =  new File(dir, fileName + "." + ext);
     if (file.exists()) {
    	return getNewFile1(file); 
     } else {
    	 return file;
     }
    
    }
    
    public static File getNewFile1(File _file) {
      	
      File rtnFile = null;
      int counter = 1;
      
      String dir = _file.getParent();
      String ext = getExtension(_file);
      String _fileName  = _file.getName();
      _fileName = _fileName.substring(0, _fileName.lastIndexOf("."));
      if (_fileName.endsWith(")")) {
        int i = _fileName.lastIndexOf("(");
        try {        
          String str = _fileName.substring(i+1,_fileName.length()-1);
          counter = Integer.parseInt(str)+1;
          _fileName = _fileName.substring(0,i);
        } catch (Exception e) {}
      }

      while (true) {
        rtnFile = new File(dir, _fileName + "(" + counter + ")." + ext);
        if (! rtnFile.exists()) {
          break;
        } else {
          counter++;
        }
      }
      return rtnFile;
    }
    
    public static void showOperationDetails() {
    	System.out.println(EXZParams.instance().getValue(OPERATION_MODE));
    }
    
}
 