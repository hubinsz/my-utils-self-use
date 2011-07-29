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
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
		 * 1 Object xlsApp, 2 Dataset ds, 3 String dotDot, 4 String period, 5
		 * String unit, 6 String title
		 */

		// 1.
		Date date = new Date();
		String fullFilePath = "fls/zzbook" + date.getTime() + ".xls";
		Workbook wb = null;
		wb = getWookBook(fullFilePath);

		// 2.
		//MyDataSet myDataSet = myDataSet = getDataSet("select TOP 10 * from pr_ReportData");
//		MyDataSet myDataSet = myDataSet = getDataSet("SELECT TOP 1000 [ID]      ,[strGUID]      ,[Code]      ,[DataType]      ,[FileGUID]      ,[PeriodID]      ,[ProjectName]      ,[Unit]      ,[GuideName]      ,[GuideUnit]      ,[Guide]      ,[Total]      ,[YearCoefficient]   FROM [quickclick].[dbo].[pr_ReportData]");
		MyDataSet myDataSet = myDataSet = getDataSet("SELECT TOP 1000 [ID]           ,[Guide]      ,[Total]      ,[YearCoefficient]   FROM [quickclick].[dbo].[pr_ReportData]");

		// 3.
		String dotDot = "2";
		int dotDigit = 0;
		try {
			dotDigit = Integer.parseInt(dotDot);
		} catch (Exception e) {
			dotDigit = 2; // default
		}

		// 4 5 6
		String title = "损益表(各种产品汇总表)";
		String unit = "单位：万元";
		String period = "第1期";

		String subTitle = period + " " + unit;

		// create sheet
		String sheetName = WorkbookUtil.createSafeSheetName(title + " "
				+ period);
		Sheet sheet = wb.createSheet(sheetName);

		// export data
		int recordCount = myDataSet.getRecordCount();
		logger.info("recordCount:" + recordCount);

		HashMap<String, Object> record = new HashMap<String, Object>();
		ArrayList<String> labelList = myDataSet.getLableList();

		// 大标题行
		sheet.addMergedRegion(new CellRangeAddress(0, // first row (0-based)
				0, // last row (0-based)
				0, // first column (0-based)
				labelList.size() - 1 // last column (0-based)
		));

		// 副标题行
		sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
				1, // last row (0-based)
				0, // first column (0-based)
				labelList.size() - 1 // last column (0-based)
		));

		// 设置大标题字体
		Font fontTitle = wb.createFont();
		fontTitle.setFontHeightInPoints((short) 22);
		fontTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// 设置副标题字体
		Font fontSubTitle = wb.createFont();
		fontSubTitle.setFontHeightInPoints((short) 11);
		fontSubTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);
		fontSubTitle.setItalic(true);

		// 设置列标题粗字体
		Font fontColumnTitle = wb.createFont();
		fontColumnTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// 在循环外面设置样式，效率比在循环里判断再修改样式高很多
		CellStyle dotStyle = wb.createCellStyle();
		CellStyle commonStringStyle = wb.createCellStyle();
		CellStyle idStyle = wb.createCellStyle();
		CellStyle colTitleStyle = wb.createCellStyle();
		CellStyle titleStyle = wb.createCellStyle();
		CellStyle subTitleStyle = wb.createCellStyle();

		colTitleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		colTitleStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		colTitleStyle.setBorderLeft(CellStyle.BORDER_THIN);
		colTitleStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		colTitleStyle.setBorderRight(CellStyle.BORDER_THIN);
		colTitleStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		colTitleStyle.setBorderTop(CellStyle.BORDER_THIN);
		colTitleStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		colTitleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		colTitleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		colTitleStyle.setFont(fontColumnTitle);

		commonStringStyle.setBorderBottom(CellStyle.BORDER_THIN);
		commonStringStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		commonStringStyle.setBorderLeft(CellStyle.BORDER_THIN);
		commonStringStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		commonStringStyle.setBorderRight(CellStyle.BORDER_THIN);
		commonStringStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		commonStringStyle.setBorderTop(CellStyle.BORDER_THIN);
		commonStringStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

		idStyle.setBorderBottom(CellStyle.BORDER_THIN);
		idStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		idStyle.setBorderLeft(CellStyle.BORDER_THIN);
		idStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		idStyle.setBorderRight(CellStyle.BORDER_THIN);
		idStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		idStyle.setBorderTop(CellStyle.BORDER_THIN);
		idStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

		dotStyle.setBorderBottom(CellStyle.BORDER_THIN);
		dotStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		dotStyle.setBorderLeft(CellStyle.BORDER_THIN);
		dotStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		dotStyle.setBorderRight(CellStyle.BORDER_THIN);
		dotStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		dotStyle.setBorderTop(CellStyle.BORDER_THIN);
		dotStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

		DataFormat format = wb.createDataFormat();

		StringBuffer sb = new StringBuffer("###,###,##0.");
		for (int i = 0; i < dotDigit; i++) {
			sb.append("0");
		}
		dotStyle.setDataFormat(format.getFormat(sb.toString()));
		idStyle.setDataFormat(format.getFormat("0"));

		double cellDigit = 0.0;

		int rowForTitle = 3;// 这个是这是此SHEET的标题和列名要占几行
		String idLableName = "ID";// 设置列名，告诉程序此列是数字，但是不要小数点，

		// 输出大标题

		Row rowTitle = sheet.createRow(0);
		Cell titleCell = rowTitle.createCell(0);
		titleCell.setCellValue(title);
		titleStyle.setFont(fontTitle);
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		titleCell.setCellStyle(titleStyle);

		// 输出副标题
		Row rowSubTitle = sheet.createRow(1);
		Cell subTitleCell = rowSubTitle.createCell(0);
		subTitleCell.setCellValue(subTitle);
		subTitleStyle.setFont(fontSubTitle);
		subTitleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		subTitleCell.setCellStyle(subTitleStyle);

		// 输出列名
		Row colTitleRow = sheet.createRow(rowForTitle - 1);
		for (int i = 0; i < labelList.size(); i++) {
			Cell cell = colTitleRow.createCell(i);
			cell.setCellValue(labelList.get(i).toString().trim());
			cell.setCellStyle(colTitleStyle);
		}

		// 输出数据
		if (recordCount > 0) {
			for (int i = rowForTitle; i < recordCount + rowForTitle; i++) {
				record = myDataSet.getRecord(i - rowForTitle);
				Row row = sheet.createRow(i);
				for (int j = 0; j < labelList.size(); j++) {
					logger.info("label is :" + labelList.get(j));
					String cellData = "";
					if (null == record.get(labelList.get(j))) {
						// do nothing
					} else {
						cellData = record.get(labelList.get(j)).toString()
								.trim();
					}
					if (isNumeric(cellData)) {
						if (idLableName.equalsIgnoreCase(labelList.get(j))) {
							Cell cell = row.createCell(j);
							cellDigit = Double.parseDouble(cellData);
							cell.setCellValue(cellDigit);
							cell.setCellStyle(idStyle);
						} else {
							logger.info("lable [" + labelList.get(j)
									+ "] is number");
							Cell cell = row.createCell(j);
							cellDigit = Double.parseDouble(cellData);
							cell.setCellValue(cellDigit);
							cell.setCellStyle(dotStyle);
						}
					} else {
						logger.info("lable [" + labelList.get(j)
								+ "] is string");
						Cell cell = row.createCell(j);
						cell.setCellValue(cellData);
						cell.setCellStyle(commonStringStyle);
					}
				}
			}
		}

		for (int i = 0; i < labelList.size(); i++) {
			sheet.autoSizeColumn(i);
		}

		// 中文的自动调整POI没有做好，还要自己加一点
		for (int i = 0; i < labelList.size(); i++) {
			sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 1000);
			if (sheet.getColumnWidth(i) < 2048) {
				sheet.setColumnWidth(i, 2048);
			}
		}

		// 如果自动调整调的列宽小于默认宽带，设回默认宽度
		for (int i = 0; i < labelList.size(); i++) {
			if (sheet.getColumnWidth(i) < 2048) {
				sheet.setColumnWidth(i, 2048);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(fullFilePath);
		wb.write(fileOut);
		fileOut.close();
		
		refinePrintSettingForWorkbook(fullFilePath);
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

	@Test
	public void otherTst() throws Exception {
		logger.info("-1   ".matches("-?\\d+(.\\d+)?"));
	}

	public Map<String, Integer> getSheetHeightAndWidth(Sheet sheet)
			throws Exception {
		int height = 0;
		int width = 0;

		int totalCellNum = 0;

		for (Row row : sheet) {
			height += row.getHeight();
		}

		Row row = sheet.getRow(sheet.getLastRowNum());
		if (row != null) {
			totalCellNum = row.getPhysicalNumberOfCells();
			for (int i = 0; i < totalCellNum; i++) {
				width += sheet.getColumnWidth(i);
			}
		}

		Map<String, Integer> retMap = new HashMap<String, Integer>();
		retMap.put("height", height);
		retMap.put("width", width);
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Sheet Name ["
				+ sheet.getSheetName() + "]");
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Height ["
				+ height + "]");
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Width ["
				+ width + "]");
		return retMap;
	}
	
	public void refinePrintSettingForOneSheet(Sheet sheet,int i) throws Exception {
		
		PrintSetup ps = sheet.getPrintSetup();

		Map<String, Integer> sheetDistanceMap = getSheetHeightAndWidth(sheet);
		int height = sheetDistanceMap.get("height");
		int width = sheetDistanceMap.get("width");
		
		ps.setFooterMargin(0.0);
		ps.setHeaderMargin(0.0);

		sheet.setMargin(Sheet.TopMargin, 0.0);
		sheet.setMargin(Sheet.BottomMargin, 0.0);
		sheet.setMargin(Sheet.LeftMargin, 0.0);
		sheet.setMargin(Sheet.RightMargin, 0.0);
		
		sheet.getWorkbook().setRepeatingRowsAndColumns(i, -1, -1, 0, 2);

		
		sheet.setAutobreaks(true);
		sheet.setHorizontallyCenter(true);
		sheet.setVerticallyCenter(true);
		ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
		
		if (height < 17000 && width < 23000) {
			ps.setFitHeight((short) 1);
			ps.setFitWidth((short) 1);
			logger.info("===**refinePrintSettingForOneSheet*=== one page, [A4], Portrait.");
		} else if (width >= 23000 && width<45000) {
			ps.setLandscape(true);
			ps.setFitWidth((short) 1);
			if (height < 9000) {
				ps.setFitHeight((short) 1);
				logger.info("===**refinePrintSettingForOneSheet*=== one page, [A4] Landscape, HorizontallyCenter, VerticallyCenter.");
			} else {
				sheet.setVerticallyCenter(false);
				ps.setFitHeight((short) 0);
				logger.info("===**refinePrintSettingForOneSheet*=== more than one page [A4], Landscape, HorizontallyCenter.");
			}
		} else {
			ps.setPaperSize(PrintSetup.A3_PAPERSIZE);
			ps.setLandscape(true);
			ps.setFitWidth((short) 1);
			if (height < 18000) {
				ps.setFitHeight((short) 1);
				logger.info("===**refinePrintSettingForOneSheet*=== one page, [A3] Landscape, HorizontallyCenter, VerticallyCenter.");
			} else {
				sheet.setVerticallyCenter(false);
				ps.setFitHeight((short) 0);
				logger.info("===**refinePrintSettingForOneSheet*=== more than one page, [A3] Landscape, HorizontallyCenter.");
			}
		}
		
		//comment next line if need print all in center
		sheet.setVerticallyCenter(false);
		
		
		sheet.setMargin(Sheet.TopMargin, 0.6);
		sheet.setMargin(Sheet.BottomMargin, 0.6);
		sheet.setMargin(Sheet.LeftMargin, 0.5);
		sheet.setMargin(Sheet.RightMargin, 0.5);

		ps.setFooterMargin(0.3);
		ps.setHeaderMargin(0.3);
		
		Footer footer = sheet.getFooter();
		Header header = sheet.getHeader();
		String sheetName = sheet.getSheetName();
		header.setRight(HeaderFooter.fontSize((short)10) + sheetName + "  Page " + HeaderFooter.page()
				+ " of " + HeaderFooter.numPages());
		footer.setRight(HeaderFooter.fontSize((short)10) + "Create Date:"+ new Date()+ "[need format!]");
	}

	public void refinePrintSettingForWorkbook(String filePath) throws Exception {
		Workbook wb = new HSSFWorkbook(new FileInputStream(
				filePath));
		int sheetNum = wb.getNumberOfSheets();
		for (int i = 0; i < sheetNum; i++) {
			refinePrintSettingForOneSheet(wb.getSheetAt(i),i);
			logger.info("======================one sheet processed======================");
		}
		FileOutputStream fileOut = new FileOutputStream(filePath);
		wb.write(fileOut);
		fileOut.close();
	}


}
