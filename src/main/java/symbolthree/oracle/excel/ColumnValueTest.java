package symbolthree.oracle.excel;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;

public class ColumnValueTest {

	
	public ColumnValueTest() {
	}

	public static void main(String[] args) {
		ColumnValueTest t = new ColumnValueTest();
		t.run();
	}
	
	private void run() {
		String jdbcurl = "jdbc:oracle:thin:@192.168.1.69:1521:APEX";
		try {
		  Connection conn = DBConnection.getInstance(jdbcurl, "system", "manager", false).getConnection();
		  Statement stmt = conn.createStatement();
		  //ResultSet rs = stmt.executeQuery("SELECT integer_col, number_col FROM columnTest where id=1");
		  ResultSet rs = stmt.executeQuery("SELECT * FROM v$session");		  
		  ResultSetMetaData rsm = rs.getMetaData();
		  for (int i=1; i<= rsm.getColumnCount(); i++) {
		    System.out.println(rsm.getColumnName(i) + " - " + rsm.getColumnTypeName(i) + " Precision = " + rsm.getPrecision(i) + ", scale = " + rsm.getScale(i));
		  }
		  rs.next();
		  double value = rs.getDouble(1);
		  if (value==0.0d) {
			  System.out.println("[]");
		  } else {
			  System.out.println("[" + rs.getDouble(1) + "]");
		  }

		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

}
