package com.sea.quickclick.report.excel.helper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
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
	public void exportExcel() throws Exception {
		/*
		 * 1 Object xlsApp, 
		 * 2 Dataset ds, 
		 * 3 String dotDot, 
		 * 4 String period, 
		 * 5 String unit,
		 * 6 String title
		 */
		
		//1. 
		String fullFilePath = "fls/Book2.xls";
		Workbook wb = null;
		wb = getWookBook(fullFilePath);
		
		//2.
		MyDataSet myDataSet = myDataSet = getDataSet("select TOP 10 * from pr_ReportData");
		
		//3.
		String dotDot = "2";
		int dotDigit = 0;
		try{
			dotDigit = Integer.parseInt(dotDot);
		}catch (Exception e) {
			dotDigit = 2; //default
		}
		
		//4 5 6
		String title = "损益表(各种产品汇总表)";
		String unit = "单位：万元";
		String period = "第1期";
		
		String subTitle = period + " " + unit;
		
		//create sheet
		String sheetName = WorkbookUtil.createSafeSheetName(title + " " + period); 
		Sheet sheet = wb.createSheet(sheetName);

	}

	private Workbook getWookBook(String fullFilePath) throws Exception {
		Workbook wb = null;
		File file = new File(fullFilePath);
		if (file.exists()) {
			// use existing file
			wb = new HSSFWorkbook(new FileInputStream(fullFilePath));
		} else {
			// create new
			wb = new HSSFWorkbook();
			// FileOutputStream fileOutputStream = new FileOutputStream(
			// "fls/workbook.xls");
			// wb.write(fileOutputStream);
			// fileOutputStream.close();
		}
		return wb;
	}

	@Test
	public void tstGetDataSet() throws Exception {
		MyDataSet myDataSet = null;
		myDataSet = getDataSet("select TOP 10 * from pr_ReportData");

		int recordCount = myDataSet.getRecordCount();
		logger.info("recordCount:" + recordCount);

		HashMap<String, Object> record = new HashMap<String, Object>();
		ArrayList<String> labelList = myDataSet.getLableList();

		if (recordCount > 0) {
			for (int i = 0; i < recordCount; i++) {
				record = myDataSet.getRecord(i);

				for (int j = 0; j < labelList.size(); j++) {
					System.out.print(record.get(labelList.get(j)) + "\t");
				}
				System.out.println();

			}

		}

	}

	public MyDataSet getDataSet(String sql) throws Exception {

		MyDataSet myDataSet = new MyDataSet();

		String url = "jdbc:sqlserver://localhost:1433;databaseName=quickclick;SelectMethod=cursor;";
		String username = "sa";
		String password = "passw0rd";
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		Connection con = DriverManager.getConnection(url, username, password);
		logger.info("connection created!");
		Statement stat = con.createStatement();
		ResultSet result = stat.executeQuery(sql);
		ResultSetMetaData resultSetMetaData = result.getMetaData();
		int columnsCount = resultSetMetaData.getColumnCount();

		ArrayList<String> labelList = new ArrayList<String>();
		List<HashMap<String, Object>> dataList = new ArrayList<HashMap<String, Object>>();

		// store labels
		for (int i = 0; i < columnsCount; i++) {
			labelList.add(i, resultSetMetaData.getColumnLabel(i + 1));
		}
		while (result.next()) {
			HashMap<String, Object> record = new HashMap<String, Object>();
			for (int i = 0; i < labelList.size(); i++) {
				record.put(labelList.get(i), result.getObject(labelList.get(i)));
			}
			dataList.add(record);
		}
		myDataSet.setLableList(labelList);
		myDataSet.setDataList(dataList);
		return myDataSet;
	}

	// internal call
	public boolean isNumeric(String str) {
		return (str.matches("-?\\d+(.\\d+)?"));
	}

}
