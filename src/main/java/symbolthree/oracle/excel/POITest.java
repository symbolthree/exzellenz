package symbolthree.oracle.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POITest {

	public POITest() {
	}

	public static void main(String[] args) {
		POITest t = new POITest();
		try {
		t.doTest();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void doTest() throws Exception {
        FileInputStream fis = new FileInputStream("C:\\Users\\Administrator\\Desktop\\test1.xls");
        Workbook        wb  = new HSSFWorkbook(fis);
        Sheet newSheet = wb.createSheet("Sheet2");
        Row row = newSheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("testing 123");
        wb.setSheetOrder("Sheet2", 0);
 
        //wb.setSheetOrder("Sheet1", 1);
        Sheet sheet = wb.cloneSheet(1);
        
        FileOutputStream fos = new FileOutputStream("C:\\\\Users\\\\Administrator\\\\Desktop\\\\test2.xls");
        wb.write(fos);
        fos.close();
		
	}

}
