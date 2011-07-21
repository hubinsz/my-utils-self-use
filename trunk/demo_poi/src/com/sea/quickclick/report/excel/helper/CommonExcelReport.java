package com.sea.quickclick.report.excel.helper;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

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
		String url = "jdbc:sqlserver://localhost:1433;databaseName=quickclick;SelectMethod=cursor;";
		String username = "sa";
		String password = "passw0rd";
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		Connection con = DriverManager.getConnection(url, username, password);
		logger.info("connection created!");
		Statement stat = con.createStatement();
		ResultSet result = stat
				.executeQuery("select TOP 10 * from pr_ReportData");
		while(result.next()){
			logger.info(result.getObject(2));
		}
		 

	}

}
