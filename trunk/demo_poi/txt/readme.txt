1.
解决了数据,时间显示格式问题
思路是把所有的数据转化成string格式,用CreationHelper这个类,
这个类可以把1.2000原样写入,后面的0不会丢失,给什么输出什么.

sheet.autoSizeColumn((short)2 );
上面行代码是自动设置列宽的最佳值.

可以跑com.sea.quickclick.report.CreateExcelDemo看效果.

2.
打印问题.
com.sea.quickclick.report.RefinePrinterSettingForWorkBook处理了.
这个类可以处理任何EXCEL,每张SHEET都会处理.
加了页眉页脚,设置最佳方式打印方式.
可以在Sample.xls添加几张SHEET,内容随便填,
先预览默认的打印效果,
再跑一下这个类,会生成Sample-print-set-up.xls文件,
再预览打印效果.

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
