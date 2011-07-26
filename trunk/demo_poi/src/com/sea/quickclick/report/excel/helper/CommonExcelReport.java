package com.sea.quickclick.report.excel.helper;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.NumberFormat;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import com.mxy.hb.poi.demo.PoiDemo001;

public class CommonExcelReport {
	static Logger logger = null;

	@BeforeClass
	public static void init() {
		logger = Logger.getLogger(PoiDemo001.class);
	}

	@AfterClass
	public static void over() {
		logger = null;
	}

	@Test
	public void getDataSet() throws Exception {
		
		MyDataSet myDataSet = new MyDataSet();
		
		
		String url = "jdbc:sqlserver://localhost:1433;databaseName=quickclick;SelectMethod=cursor;";
		String username = "sa";
		String password = "passw0rd";
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		Connection con = DriverManager.getConnection(url, username, password);
		logger.info("connection created!");
		Statement stat = con.createStatement();
		ResultSet result = stat
				.executeQuery("select TOP 10 * from pr_ReportData");
		ResultSetMetaData resultSetMetaData = result.getMetaData();
		int columnsCount = resultSetMetaData.getColumnCount();
		String[] columnNames = new String[columnsCount];
		//ArrayList<String>
		for (int i = 0; i < columnsCount; i++) {
			columnNames[i] = resultSetMetaData.getColumnLabel(i + 1);
		}
		
		for(int i=0;i<columnNames.length;i++){
			logger.info(columnNames[i]);
		}
		
		while (result.next()) {
		//	logger.info(result.getMetaData());
		}

	}

	@Test
	public void isNumeric() {

		logger.info("2344234.33".matches("-?\\d+(.\\d+)?"));

	}

}
