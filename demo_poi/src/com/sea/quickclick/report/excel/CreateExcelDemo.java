package com.sea.quickclick.report.excel;


import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CreateExcelDemo {

	public static void main(String[] args) throws Exception{
			Workbook wb = new HSSFWorkbook();
			CreationHelper createHelper = wb.getCreationHelper();
			Sheet sheet = wb.createSheet("sheet1");
			Row row = sheet.createRow((short) 3);
			
			Cell cell = row.createCell(0);
			cell.setCellValue(1.00);
			row.createCell(1).setCellValue(1.20);
			row.createCell(2).setCellValue(
					createHelper.createRichTextString("test rich string!"));
			row.createCell(3).setCellValue(createHelper.createRichTextString("1.20000"));
			row.createCell(4).setCellValue(createHelper.createRichTextString("2011-07-16 12:12"));

			//�����Զ�������ѿ��
			sheet.autoSizeColumn((short)2 );
			sheet.autoSizeColumn((short)3 );
			sheet.autoSizeColumn((short)4 );

			FileOutputStream fileOut = new FileOutputStream("fls/demo_text_out.xls");
			wb.write(fileOut);
			fileOut.close();

		}

	}


