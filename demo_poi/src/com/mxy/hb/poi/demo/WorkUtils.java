package com.mxy.hb.poi.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
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
				"fls/Sample.xls"));
		int sheetNum = wb.getNumberOfSheets();
		for (int i = 0; i < sheetNum; i++) {
			refinePrintSettingForOneSheet(wb.getSheetAt(i));
			logger.info("======================one sheet processed======================");
		}
		FileOutputStream fileOut = new FileOutputStream("fls/Sample-del.xls");
		wb.write(fileOut);
		fileOut.close();
	}

	public void refinePrintSettingForOneSheet(Sheet sheet) throws Exception {
		
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
		
		sheet.setAutobreaks(true);
		sheet.setHorizontallyCenter(true);
		sheet.setVerticallyCenter(true);

		if (height < 17000 && width < 21000) {
			ps.setFitHeight((short) 1);
			ps.setFitWidth((short) 1);
			logger.info("===**refinePrintSettingForOneSheet*=== one page, Portrait.");
		} else if (width >= 21000 && width<45000) {
			ps.setLandscape(true);
			ps.setFitWidth((short) 1);
			if (height < 9000) {
				ps.setFitHeight((short) 1);
				logger.info("===**refinePrintSettingForOneSheet*=== one page, Landscape, HorizontallyCenter, VerticallyCenter.");
			} else {
				sheet.setVerticallyCenter(false);
				ps.setFitHeight((short) 0);
				logger.info("===**refinePrintSettingForOneSheet*=== more than one page, Landscape, HorizontallyCenter.");
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

}
