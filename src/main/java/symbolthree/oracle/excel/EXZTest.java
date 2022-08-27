package symbolthree.oracle.excel;

//~--- non-JDK imports --------------------------------------------------------

import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class EXZTest {
    public static void main(String[] args) {
        EXZTest test = new EXZTest();

        try {
            //System.out.println(EXZHelper.getNewFile(new File("C:\\sqlnet(wed).log")));
            //test.start();
        	test.newFile("abcde_$D{YYYY-xx-dd}");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    
    private void newFile(String _fileName) {
    	
  	  Date currDate = new Date();
  	  String output = null;
  	  
  	  System.out.println("input = " + _fileName);
  	  
  	  int pos1 = _fileName.indexOf("$D{");
  	  System.out.println("pos1=" + pos1);
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
  		System.out.println("timestamp = "  + timestamp);
  		String result = timestamp;
  		try {
    	SimpleDateFormat format = new SimpleDateFormat(timestamp);
    	result = format.format(currDate);
    	System.out.println("timestamp = "  + result);    	
  		} catch (Exception e) {
  		}

  		output = _fileName.substring(0, pos1) + result + _fileName.substring(pos2+1);
  	  }
  	  
  	  System.out.println("output = " + output);
    }
    
    private void start() throws Exception {
        File            file      = new File("F:\\OracleTools\\EXZELLENZ\\Test.xls");
        FileInputStream fis       = new FileInputStream(file);
        HSSFWorkbook    wb        = new HSSFWorkbook(fis);
        HSSFSheet       sheet     = wb.getSheetAt(0);
        HSSFRow         row       = sheet.getRow(0);
        HSSFCell        cell      = row.getCell(0);
        HSSFCellStyle   cellStyle = cell.getCellStyle();

        System.out.println(EXZHelper.number2Letter(1));
        System.out.println(cellStyle.getDataFormatString(wb));
        row  = sheet.createRow(1);
        cell = row.createCell(1);

        List     list = HSSFDataFormat.getBuiltinFormats();
        Iterator itr  = list.iterator();

        while (itr.hasNext()) {
            System.out.println((String) itr.next());
        }

        /*
         * cellStyle = wb.createCellStyle();
         * cellStyle.setDataFormat()
         * cell.setCellStyle(cellStyle);
         * cell.setCellValue(Calendar.getInstance().getTime());
         */
        FileOutputStream fos = new FileOutputStream(file);

        wb.write(fos);
        fos.close();
    }
}
