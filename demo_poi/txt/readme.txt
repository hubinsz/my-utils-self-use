1.
���������,ʱ����ʾ��ʽ����
˼·�ǰ����е�����ת����string��ʽ,��CreationHelper�����,
�������԰�1.2000ԭ��д��,�����0���ᶪʧ,��ʲô���ʲô.

sheet.autoSizeColumn((short)2 );
�����д������Զ������п�����ֵ.

������com.sea.quickclick.report.CreateExcelDemo��Ч��.

2.
��ӡ����.
com.sea.quickclick.report.RefinePrinterSettingForWorkBook������.
�������Դ����κ�EXCEL,ÿ��SHEET���ᴦ��.
����ҳüҳ��,������ѷ�ʽ��ӡ��ʽ.
������Sample.xls��Ӽ���SHEET,���������,
��Ԥ��Ĭ�ϵĴ�ӡЧ��,
����һ�������,������Sample-print-set-up.xls�ļ�,
��Ԥ����ӡЧ��.

=============
package com.sea.quickclick.report.excel.helper;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import com.sea.quickclick.data.Dataset;
import com.sea.quickclick.helper.DatasetHelper;

public class CommonExcelReport {
	static Logger logger = null;

	@BeforeClass
	public static void init() {
		logger = Logger.getRootLogger();
	}

	@AfterClass
	public static void over() {
		logger = null;
	}

	public void exportExcel() throws Exception {
		// Dataset
	}

	public boolean exportExcel(String filePath, Dataset ds, String dotDot,
			String period, String unit, String title) throws Exception {

		
		
		
		
		return true;
	}

	@Test
	public void createExcel() throws Exception {
		String filePath = "fls/createByApp.xls";
		Workbook wb = new HSSFWorkbook();
		Sheet sheet1 = wb.createSheet("sheet1");

		for (int r = 2; r < 150; r++) {
			Row row = sheet1.createRow(r);
			Cell cell = null;
			for (int c = 0; c < 60; c++) {
				cell = row.createCell(c);
				cell.setCellValue("r_" + r + "c_" + c);
			}
		}


		
		File file = new File(filePath);
		if (file.exists()) {
			logger.info("["+filePath+"] already existing!");
			logger.info("["+filePath+"] absolutePath is: "+file.getAbsolutePath());
			file.delete();
			logger.info("["+filePath+"] deleted by app!");
		}

		FileOutputStream fileOutputStream = new FileOutputStream(filePath);
		wb.write(fileOutputStream);
		fileOutputStream.close();
		logger.info("[" + filePath + "] has been created!");

	}

}
