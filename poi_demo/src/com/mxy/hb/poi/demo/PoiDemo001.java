package com.mxy.hb.poi.demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

public class PoiDemo001 {

	static Logger logger = null;

	@BeforeClass
	public static void init() {
		logger = Logger.getLogger(PoiDemo001.class);
	}

	@AfterClass
	public static void over() {
		logger = null;
	}

	// How to create a new workbook
	// only a excel file, no data yet
	@Test
	public void createWorkbook001() throws Exception {
		Workbook wb = new HSSFWorkbook();
		FileOutputStream fileOutputStream = new FileOutputStream(
				"fls/workbook.xls");
		wb.write(fileOutputStream);
		fileOutputStream.close();
	}

	// How to create a sheet
	@Test
	public void createSheet002() throws Exception {
		Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
		Sheet sheet1 = wb.createSheet("new sheet");
		Sheet sheet2 = wb.createSheet("second sheet");

		// Note that sheet name is Excel must not exceed 31 characters
		// and must not contain any of the any of the following characters:
		// 0x0000
		// 0x0003
		// colon (:)
		// backslash (\)
		// asterisk (*)
		// question mark (?)
		// forward slash (/)
		// opening square bracket ([)
		// closing square bracket (])

		// You can use
		// org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String
		// nameProposal)}
		// for a safe way to create valid names, this utility replaces invalid
		// characters with a space (' ')
		String safeName = WorkbookUtil
				.createSafeSheetName("[O'Brien's sales*?]"); // returns
																// " O'Brien's sales   "
		Sheet sheet3 = wb.createSheet(safeName);

		FileOutputStream fileOut = new FileOutputStream("fls/workbook.xls");
		wb.write(fileOut);
		fileOut.close();
	}

	// How to create cells
	@Test
	public void createCells003() throws Exception {
		Workbook wb = new HSSFWorkbook();
		// Workbook wb = new XSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");
		// add by hubin, test column width
		sheet.setColumnWidth(1, 5000);

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow((short) 5);
		// Create a cell and put a value in it.
		Cell cell = row.createCell(0);
		cell.setCellValue(1);

		// Or do it on one line.
		row.createCell(1).setCellValue(1.20);
		row.createCell(2).setCellValue(
				createHelper.createRichTextString("This is a string"));
		row.createCell(3).setCellValue(true);
		row.createCell(4).setCellValue(createHelper.createRichTextString("1.20"));
		row.createCell(5).setCellValue(createHelper.createRichTextString("2011-07-16"));

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook003.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// How to create date cells
	@Test
	public void createDateCell004() throws Exception {
		Workbook wb = new HSSFWorkbook();
		// Workbook wb = new XSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow(0);

		// Create a cell and put a date value in it. The first cell is not
		// styled
		// as a date.
		Cell cell = row.createCell(0);
		cell.setCellValue(new Date());

		// we style the second cell as a date (and time). It is important to
		// create a new cell style from the workbook otherwise you can end up
		// modifying the built in style and effecting not only this cell but
		// other cells.
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"m/d/yy h:mm"));
		cell = row.createCell(1);
		cell.setCellValue(new Date());
		cell.setCellStyle(cellStyle);

		// you can also set date as java.util.Calendar
		cell = row.createCell(2);
		cell.setCellValue(Calendar.getInstance());
		cell.setCellStyle(cellStyle);

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook004.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Working with different types of cells
	@Test
	public void createDiffTypeCell005() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");
		Row row = sheet.createRow((short) 2);
		row.createCell(0).setCellValue(1.1);
		row.createCell(1).setCellValue(new Date());
		row.createCell(2).setCellValue(Calendar.getInstance());
		row.createCell(3).setCellValue("a string");
		row.createCell(4).setCellValue(true);
		row.createCell(5).setCellType(Cell.CELL_TYPE_ERROR);

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook005.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Demonstrates various alignment options
	@Test
	public void createAlignment006() throws Exception {
		Workbook wb = new HSSFWorkbook(); // or new HSSFWorkbook();

		Sheet sheet = wb.createSheet("test");
		// add by me to set cell width
		// sheet.autoSizeColumn((short) 0); did not work
		// sheet.setColumnWidth(0, 5000);
		Row row = sheet.createRow((short) 2);
		row.setHeightInPoints(30);

		createCell(wb, row, (short) 0, CellStyle.ALIGN_CENTER,
				CellStyle.VERTICAL_BOTTOM);
		createCell(wb, row, (short) 1, CellStyle.ALIGN_CENTER_SELECTION,
				CellStyle.VERTICAL_BOTTOM);
		createCell(wb, row, (short) 2, CellStyle.ALIGN_FILL,
				CellStyle.VERTICAL_CENTER);
		createCell(wb, row, (short) 3, CellStyle.ALIGN_GENERAL,
				CellStyle.VERTICAL_CENTER);
		createCell(wb, row, (short) 4, CellStyle.ALIGN_JUSTIFY,
				CellStyle.VERTICAL_JUSTIFY);
		createCell(wb, row, (short) 5, CellStyle.ALIGN_LEFT,
				CellStyle.VERTICAL_TOP);
		createCell(wb, row, (short) 6, CellStyle.ALIGN_RIGHT,
				CellStyle.VERTICAL_TOP);

		sheet.autoSizeColumn((short) 2);
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook006.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	/**
	 * Creates a cell and aligns it a certain way.
	 * 
	 * @param wb
	 *            the workbook
	 * @param row
	 *            the row to create the cell in
	 * @param column
	 *            the column number to create the cell in
	 * @param halign
	 *            the horizontal alignment for the cell.
	 */
	private void createCell(Workbook wb, Row row, short column, short halign,
			short valign) {
		Cell cell = row.createCell(column);
		cell.setCellValue("Align ItAlign ItAlign ItAlign ItAlign It");
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(valign);
		cell.setCellStyle(cellStyle);
	}

	// Working with borders
	@Test
	public void createBorders007() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow(1);

		// Create a cell and put a value in it.
		Cell cell = row.createCell(1);
		cell.setCellValue(4);

		// Style the cell with borders all around.
		CellStyle style = wb.createCellStyle();
		//
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLUE.getIndex());
		style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		//
		cell.setCellStyle(style);

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook007.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Iterate over rows and cells using Java 1.5 foreach loops
	@Test
	public void showSheet008() throws Exception {
		Workbook wb = new HSSFWorkbook(new FileInputStream("fls/dddd.xls"));
		Sheet sheet = wb.getSheetAt(0);
		int i = 0;
		for (Row row : sheet) {
			for (Cell cell : row) {
				logger.debug("[" + (i++) + "]" + cell.getStringCellValue());
			}
		}
	}

	// Getting the cell contents
	@Test
	public void fetchAllData009() throws Exception {
		Workbook wb = new HSSFWorkbook(new FileInputStream("fls/dddd.xls"));
		Sheet sheet1 = wb.getSheetAt(0);
		int i = 0;
		for (Row row : sheet1) {
			logger.debug(i++);
			for (Cell cell : row) {
				CellReference cellRef = new CellReference(row.getRowNum(),
						cell.getColumnIndex());
				System.out.print(cellRef.formatAsString());
				System.out.print(" - ");

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getRichStringCellValue()
							.getString());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						System.out.println(cell.getDateCellValue());
					} else {
						System.out.println(cell.getNumericCellValue());
					}
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					System.out.println(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					System.out.println(cell.getCellFormula());
					break;
				default:
					System.out.println();
				}
			}
		}

	}

	// Text Extraction
	@Test
	public void textExtraFromXls010() throws Exception {
		InputStream inp = new FileInputStream("fls/dddd.xls");
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
		ExcelExtractor extractor = new ExcelExtractor(wb);

		extractor.setFormulasNotResults(true);
		extractor.setIncludeSheetNames(false);
		String text = extractor.getText();
		logger.debug(text);
	}
	
	// re-set printer
	@Test
	public void resetPrinter() throws Exception {
		InputStream inp = new FileInputStream("fls/workbook100.xls");
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
		Sheet sheet = wb.getSheetAt(0);
		PrintSetup ps = sheet.getPrintSetup();
		ps.setLandscape(true);
		sheet.setVerticallyCenter(true);
		sheet.setHorizontallyCenter(true);

		
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook011.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Fills and colors
	@Test
	public void createFillsAndColors011() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow((short) 1);

		// Aqua background
		CellStyle style = wb.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
		style.setFillPattern(CellStyle.BIG_SPOTS);
		Cell cell = row.createCell((short) 1);
		cell.setCellValue("X");
		cell.setCellStyle(style);

		// Orange "foreground", foreground being the fill foreground not the
		// font color.
		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cell = row.createCell((short) 2);
		cell.setCellValue("X");
		cell.setCellStyle(style);

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook011.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Merging cells
	@Test
	public void mergeCells012() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");

		Row row = sheet.createRow((short) 1);
		Cell cell = row.createCell((short) 1);
		cell.setCellValue("This is a test of merging");

		sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
				1, // last row (0-based)
				1, // first column (0-based)
				2 // last column (0-based)
		));

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook012.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Working with fonts
	@Test
	public void testFonts013() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");

		// Create a row and put some cells in it. Rows are 0 based.
		Row row = sheet.createRow(1);

		// Create a new font and alter it.
		Font font = wb.createFont();
		font.setFontHeightInPoints((short) 24);
		font.setFontName("Courier New");
		font.setItalic(true);
		font.setStrikeout(true);

		// Fonts are set into a style so create a new one to use.
		CellStyle style = wb.createCellStyle();
		style.setFont(font);

		// Create a cell and put a value in it.
		Cell cell = row.createCell(1);
		cell.setCellValue("This is a test of fonts");
		cell.setCellStyle(style);

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook013.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	/*
	 * Note, the maximum number of unique fonts in a workbook is limited to
	 * 32767 ( the maximum positive short). You should re-use fonts in your
	 * apllications instead of creating a font for each cell. Examples:
	 * 
	 * Wrong:
	 * 
	 * for (int i = 0; i < 10000; i++) { Row row = sheet.createRow(i); Cell cell
	 * = row.createCell((short) 0);
	 * 
	 * CellStyle style = workbook.createCellStyle(); Font font =
	 * workbook.createFont(); font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	 * style.setFont(font); cell.setCellStyle(style); } Correct:
	 * 
	 * 
	 * CellStyle style = workbook.createCellStyle(); Font font =
	 * workbook.createFont(); font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	 * style.setFont(font); for (int i = 0; i < 10000; i++) { Row row =
	 * sheet.createRow(i); Cell cell = row.createCell((short) 0);
	 * cell.setCellStyle(style); }
	 */

	// Custom colors
	@Test
	public void customColor014() throws Exception {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet();
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCell cell = row.createCell((short) 0);
		cell.setCellValue("Default Palette");

		// apply some colors from the standard palette,
		// as in the previous examples.
		// we'll use red text on a lime background

		HSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(HSSFColor.LIME.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		HSSFFont font = wb.createFont();
		font.setColor(HSSFColor.RED.index);
		style.setFont(font);

		cell.setCellStyle(style);

		// save with the default palette
		FileOutputStream out = new FileOutputStream("fls/default_palette.xls");
		wb.write(out);
		out.close();

		// now, let's replace RED and LIME in the palette
		// with a more attractive combination
		// (lovingly borrowed from freebsd.org)

		cell.setCellValue("Modified Palette");

		// creating a custom palette for the workbook
		HSSFPalette palette = wb.getCustomPalette();

		// replacing the standard red with freebsd.org red
		palette.setColorAtIndex(HSSFColor.RED.index, (byte) 153, // RGB red
																	// (0-255)
				(byte) 0, // RGB green
				(byte) 0 // RGB blue
		);
		// replacing lime with freebsd.org gold
		palette.setColorAtIndex(HSSFColor.LIME.index, (byte) 255, (byte) 204,
				(byte) 102);

		// save with the modified palette
		// note that wherever we have previously used RED or LIME, the
		// new colors magically appear
		out = new FileOutputStream("fls/modified_palette.xls");
		wb.write(out);
		out.close();

	}

	// Reading and Rewriting Workbooks
	@Test
	public void readAndRewrite015() throws Exception {
		InputStream inp = new FileInputStream("fls/workbook015.xls");
		// InputStream inp = new FileInputStream("workbook.xlsx");

		Workbook wb = new HSSFWorkbook(inp);
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(2);
		Cell cell = row.getCell(3);
		if (cell == null)
			cell = row.createCell(3);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("a test");

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("fls/workbook015.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Using newlines in cells
	@Test
	public void newLines016() throws Exception {
		Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
		Sheet sheet = wb.createSheet();

		Row row = sheet.createRow(2);
		Cell cell = row.createCell(2);
		cell.setCellValue("Use \n with word wrap on to create a new line");

		// to enable newlines you need set a cell styles with wrap=true
		CellStyle cs = wb.createCellStyle();
		cs.setWrapText(true);
		cell.setCellStyle(cs);

		// increase row height to accomodate two lines of text
		row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));

		// adjust column width to fit the content
		sheet.autoSizeColumn((short) 2);

		FileOutputStream fileOut = new FileOutputStream(
				"fls/ooxml-newlines.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Data Formats
	@Test
	public void dataFormat017() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("format sheet");
		CellStyle style;
		DataFormat format = wb.createDataFormat();
		Row row;
		Cell cell;
		short rowNum = 0;
		short colNum = 0;

		row = sheet.createRow(rowNum++);
		cell = row.createCell(colNum);
		cell.setCellValue(11111.25);
		style = wb.createCellStyle();
		style.setDataFormat(format.getFormat("0.0"));
		cell.setCellStyle(style);

		row = sheet.createRow(rowNum++);
		cell = row.createCell(colNum);
		cell.setCellValue(11111.25);
		style = wb.createCellStyle();
		style.setDataFormat(format.getFormat("#,##0.0000"));
		cell.setCellStyle(style);

		sheet.autoSizeColumn((short) 0);
		FileOutputStream fileOut = new FileOutputStream("fls/workbook017.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Fit Sheet to One Page
	@Test
	public void fitSheet2OnePage018() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("format sheet");
		PrintSetup ps = sheet.getPrintSetup();

		sheet.setAutobreaks(true);

		ps.setFitHeight((short) 1);
		ps.setFitWidth((short) 1);

		// Create various cells and rows for spreadsheet.
		for (int i = 0; i < 70; i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < 15; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue("hb" + j + i);
			}
		}

		FileOutputStream fileOut = new FileOutputStream("fls/workbook018.xls");
		wb.write(fileOut);
		fileOut.close();
	}
	// setLandscape
	@Test
	public void setLandscapeDemo001() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("format sheet");
		PrintSetup ps = sheet.getPrintSetup();
		
		
		sheet.setAutobreaks(true);
		ps.setFitHeight((short) 1);
		ps.setFitWidth((short) 1);
		ps.setLandscape(true);
		// Create various cells and rows for spreadsheet.
		for (int i = 0; i < 70; i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < 15; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue("hb" + j + i);
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream("fls/workbook100.xls");
		wb.write(fileOut);
		fileOut.close();
	}

	// Set Print Area
	@Test
	public void configPrintArea019() throws Exception {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		// sets the print area for the first sheet
		// wb.setPrintArea(0, "$A$1:$C$2");
		for (int i = 0; i < 70; i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < 15; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue("col" + j + ",row" + i);
			}
		}
		// Alternatively:
		wb.setPrintArea(0, // sheet index
				0, // start column
				10, // end column
				0, // start row
				10 // end row
		);

		FileOutputStream fileOut = new FileOutputStream("fls/workbook019.xls");
		wb.write(fileOut);
		fileOut.close();

	}

	// Set Page Numbers on Footer
	@Test
	public void setPageNum020() throws Exception {
		Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
		Sheet sheet = wb.createSheet("format sheet");
		Footer footer = sheet.getFooter();

		footer.setRight("Page " + HeaderFooter.page() + " of "
				+ HeaderFooter.numPages());

		// Create various cells and rows for spreadsheet.
		for (int i = 0; i < 70; i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < 15; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue("col" + j + ",row" + i);
			}
		}

		FileOutputStream fileOut = new FileOutputStream("fls/workbook020.xls");
		wb.write(fileOut);
		fileOut.close();

	}
}
