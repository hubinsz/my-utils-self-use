package com.mxy.hb.poi.demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

public class PrintTestDome {
	static Logger logger = null;

	@BeforeClass
	public static void init(){
		logger = Logger.getLogger(PoiDemo001.class);
	}
	
	@AfterClass
	public static void over(){
		logger = null;
	}

	@Test
	public void testPrint001() throws Exception{
		Workbook wb = new HSSFWorkbook(new FileInputStream("fls/testPrint.xls"));
		Sheet sheet = wb.getSheetAt(0);
		//sheet.setFitToPage(true);
		
		sheet.setPrintGridlines(false);
		sheet.setAutobreaks(true);
		sheet.setMargin(Sheet.TopMargin, 0.5);
		double i = sheet.getMargin(Sheet.LeftMargin);
		logger.debug(i);
		PrintSetup printSetup = sheet.getPrintSetup();
		//printSetup.setFitHeight(arg0)
		//printSetup.set
		printSetup.setFitWidth((short)1);
		printSetup.setFitHeight((short)10);
		double y = printSetup.getFooterMargin();
		logger.debug(y);
		//printSetup.
		
		
		
		
		sheet.setVerticallyCenter(true);
		sheet.setHorizontallyCenter(true);
		
		FileOutputStream fileOutputStream = new FileOutputStream("fls/testPrint.xls");
		wb.write(fileOutputStream);
		fileOutputStream.close();

	}
	@Test
	public void testPrint002() throws Exception{
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("sheet1");
		
		for(int i=0;i<50;i++){
			Row row = sheet.createRow(i);
			for(int j=0;j<9;j++){
				Cell cell = row.createCell(j);
				cell.setCellValue("r"+i+",c"+j);
			}
			
		}
		
		int reNum = sheet.getPhysicalNumberOfRows();
		logger.debug(reNum);
		
		Row roww = sheet.getRow(30);
		int sss = roww.getPhysicalNumberOfCells();
		logger.debug(sss);
		
		int colLength = 0;
		int rowLength = 0;
		
		for (Row row : sheet){
			rowLength +=row.getHeight();
		}
		
		Row roww2 = sheet.getRow(30);
		for(int k=0;k<sss;k++){
			colLength += sheet.getColumnWidth(k);
		}
		
		logger.debug(colLength);
		logger.debug(rowLength);
		for(Cell cell : roww2){
			
		}
		
		
        for (Row row : sheet) {
        	
            for (Cell cell : row) {

            
            }
        }

		
		/*
		//sheet.setFitToPage(true);
		
		sheet.setPrintGridlines(false);
		sheet.setAutobreaks(true);
		sheet.setMargin(Sheet.TopMargin, 0.5);
		double i = sheet.getMargin(Sheet.LeftMargin);
		logger.debug(i);
		PrintSetup printSetup = sheet.getPrintSetup();
		//printSetup.setFitHeight(arg0)
		//printSetup.set
		printSetup.setFitWidth((short)1);
		printSetup.setFitHeight((short)10);
		double y = printSetup.getFooterMargin();
		logger.debug(y);
		//printSetup.
		
		
		
		
		sheet.setVerticallyCenter(true);
		sheet.setHorizontallyCenter(true);
		*/
		FileOutputStream fileOutputStream = new FileOutputStream("fls/testPrint002.xls");
		wb.write(fileOutputStream);
		fileOutputStream.close();
		
	}
}
