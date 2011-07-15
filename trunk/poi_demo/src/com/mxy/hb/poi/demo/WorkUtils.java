package com.mxy.hb.poi.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

public class WorkUtils {
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
	public void refinePrintSettingForWorkbook() throws Exception {
		Workbook wb = new HSSFWorkbook(new FileInputStream(
		// "fls/testPrint002.xls"));
				"fls/Book1.xls"));
		int sheetNum = wb.getNumberOfSheets();
		for (int i = 0; i < sheetNum; i++) {
			refinePrintSettingForOneSheet(wb.getSheetAt(i));
			logger.info("======================one sheet processed======================");
		}
		FileOutputStream fileOut = new FileOutputStream("fls/Book2.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	public void refinePrintSettingForOneSheet(Sheet sheet) throws Exception {
		Footer footer = sheet.getFooter();
		Header header = sheet.getHeader();
		String sheetName = sheet.getSheetName();
		String footerHeaderTitle = sheetName + "  Page " + HeaderFooter.page()
				+ " of " + HeaderFooter.numPages();
		footer.setCenter(footerHeaderTitle);
		header.setCenter(footerHeaderTitle);

		

		PrintSetup ps = sheet.getPrintSetup();

		Map<String, Integer> sheetDistanceMap = getSheetHeightAndWidth(sheet);
		int height = sheetDistanceMap.get("height");
		int width = sheetDistanceMap.get("width");

		if (height < 17000 && width < 21000) {
			ps.setFitHeight((short) 1);
			ps.setFitWidth((short) 1);
			sheet.setHorizontallyCenter(true);
			sheet.setVerticallyCenter(true);
			logger.info("===**refinePrintSettingForOneSheet*=== one page, Portrait.");
		} else if (width < 30000 && width >= 21000) {
			
			ps.setLandscape(true);
			if (height < 9000) {
				sheet.setAutobreaks(true);
				ps.setFitHeight((short) 1);
				ps.setFitWidth((short) 1);
				sheet.setVerticallyCenter(true);
				sheet.setHorizontallyCenter(true);
				logger.info("===**refinePrintSettingForOneSheet*=== one page, Landscape, HorizontallyCenter, VerticallyCenter.");
			} else {
				sheet.setAutobreaks(true);
				ps.setFitHeight((short) 0);
				ps.setFitWidth((short) 1);
				sheet.setVerticallyCenter(false);
				sheet.setHorizontallyCenter(true);
				logger.info("===**refinePrintSettingForOneSheet*=== more than one page, Landscape, HorizontallyCenter.");
			}
		} else {
			sheet.setHorizontallyCenter(true);
			sheet.setVerticallyCenter(true);
			logger.info("===**refinePrintSettingForOneSheet*=== other distance case.");
		}
		/*
		 * File file = new File("fls/Book2.xls"); if(file.exists()){
		 * file.delete(); }
		 */
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
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Sheet Name [" + sheet.getSheetName() + "]");
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Height [" + height + "]");
		logger.info("===***getSheetHeightAndWidth(Sheet sheet)=== Width [" + width + "]");
		return retMap;
	}

}
