package symbolthree.oracle.excel;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;

import oracle.jdbc.OracleCallableStatement;
import oracle.jdbc.OraclePreparedStatement;
import oracle.jdbc.OracleTypes;

public class JDBCTest {

	public JDBCTest() {
	}

	public static void main(String[] args) {
		JDBCTest t= new JDBCTest();
		t.run2();
	}
	
	private void run() {
		String jdbcurl = "jdbc:oracle:thin:@192.168.1.175:1521:PS";
		try {
		  Connection conn = DBConnection.getInstance(jdbcurl, "SCOTT", "tiger", false).getConnection();
		  String sql = "INSERT INTO EMP (EMPNO, ENAME, JOB) VALUES (?, ?, ?) RETURNING ROWIDTOCHAR(ROWID) INTO ?";
		  
		  OracleCallableStatement stmt = (OracleCallableStatement)conn.prepareCall(sql);
		  
		  stmt.setInt(1, 2000);
		  stmt.setString(2, "MYNAME");
		  stmt.setString(3, "BLOWJOB");
		  stmt.registerReturnParameter(4, OracleTypes.VARCHAR);
		  stmt.executeUpdate();
		  ResultSet rs = stmt.getReturnResultSet();
		  rs.next();
		  String rowid = rs.getString(1);
		  System.out.println("ROWID=" + rowid);
		  conn.rollback();
		  stmt.close();
		  conn.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	private void run2() {
		String jdbcurl = "jdbc:oracle:thin:@192.168.1.113:1541:VCPDEMO";
		try {
		Connection conn = DBConnection.getInstance(jdbcurl, "APPS", "APPS", false).getConnection();
		String proc = "DECLARE " + 
					"BOO BOOLEAN;" + 
					"RTN VARCHAR2(1);" + 
					"BEGIN " + 
					"BOO := fnd_user_pkg.validatelogin(?,?);" + 
					"IF BOO THEN RTN := 'Y'; ELSE RTN := 'N';END IF;" + 
					"? := RTN;" + 
					"END;"; 

		CallableStatement stmt = conn.prepareCall(proc);
		stmt.setString(1, "SYSADMIN");
		stmt.setString(2, "sysadmin");
		stmt.registerOutParameter(3, java.sql.Types.VARCHAR);		
		stmt.execute();
		String boo = stmt.getString(3);
		System.out.println("User/password combination valid : " + boo);
		
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
	}

}
