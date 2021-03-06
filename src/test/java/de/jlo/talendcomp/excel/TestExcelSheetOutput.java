package de.jlo.talendcomp.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;

public class TestExcelSheetOutput {
	
	private Map<String, Object> globalMap = new HashMap<>();
	
	public org.apache.poi.ss.usermodel.Workbook setupEmptyWorkbook() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		tFileExcelWorkbookOpen_1.initializeWorkbook();
		return tFileExcelWorkbookOpen_1.getWorkbook();
	}
	
	@Test
	public void testWriteZeroDate() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		tFileExcelWorkbookOpen_1.setInputFile(
				"/var/testdata/excel/time.xlsx", false);
		tFileExcelWorkbookOpen_1.initializeWorkbook();
		
		final de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen_1.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName("out");
		tFileExcelSheetOutput_1.resetCache();
		tFileExcelSheetOutput_1.setRowStartIndex(1 - 1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(false);
		// configure cell positions
		tFileExcelSheetOutput_1.setColumnStart("A");
		tFileExcelSheetOutput_1.setReuseExistingStylesFromFirstWrittenRow(false);
		tFileExcelSheetOutput_1.setSetupCellStylesForAllColumns(false);
		tFileExcelSheetOutput_1.setReuseFirstRowHeight(false);
		// configure cell formats
		// columnIndex: 0, format: "dd.MM.yyyy", talendType: Date
		tFileExcelSheetOutput_1.setDataFormat(0, "dd.mm.yyyy");
		tFileExcelSheetOutput_1.setDataFormat(1, "m/d/yy h:mm");
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		tFileExcelSheetOutput_1.setWriteZeroDateAsNull(true);
		String dateStr = "0000-00-00";
		String dateStr2 = "2021-06-23";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		Date zeroDate = sdf.parse(dateStr);
		Object[] row = new Object[2];
		row[0] = zeroDate;
		row[1] = sdf.parse(dateStr2);
		tFileExcelSheetOutput_1.writeRow(row);
		tFileExcelSheetOutput_1.writeRow(row);

		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1.setWorkbook(tFileExcelWorkbookOpen_1.getWorkbook());
		tFileExcelWorkbookSave_1.evaluateAllFormulars();
		assertEquals(1, tFileExcelSheetOutput_1.detectCurrentSheetLastNoneEmptyRowIndex());
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_1
					.setOutputFile("/var/testdata/excel/time_out2.xlsx");
			tFileExcelWorkbookSave_1.createDirs();
			tFileExcelWorkbookSave_1.writeWorkbook();
			// release the memory
		} catch (Exception e) {
			throw e;
		}
		assertTrue(true);
	}

	@Test
	public void testOutputFormats() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_2 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_2.setCreateStreamingXMLWorkbook(false);
		try {
			// create empty XLSX workbook
			tFileExcelWorkbookOpen_2.createEmptyXLSXWorkbook();
			tFileExcelWorkbookOpen_2.initializeWorkbook();
		} catch (Exception e) {
			tFileExcelWorkbookOpen_2.error(
					"Intialize empty workbook failed: "
							+ e.getMessage(), e);
			globalMap.put("tFileExcelWorkbookOpen_2_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_2",
				tFileExcelWorkbookOpen_2.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_2_COUNT_SHEETS",
				tFileExcelWorkbookOpen_2.getWorkbook()
						.getNumberOfSheets());

		final de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_2"));
		tFileExcelSheetOutput_1.setTargetSheetName("test_out");
		globalMap.put("tFileExcelSheetOutput_1_SHEET_NAME",
				tFileExcelSheetOutput_1.getTargetSheetName());
		tFileExcelSheetOutput_1.resetCache();
		tFileExcelSheetOutput_1.setRowStartIndex(1 - 1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(true);
		// configure cell positions
		tFileExcelSheetOutput_1.setDataColumnPosition(0, "A");
		tFileExcelSheetOutput_1.setDataColumnPosition(1, "B");
		tFileExcelSheetOutput_1.setDataColumnPosition(2, "K");
		tFileExcelSheetOutput_1.setDataColumnPosition(3, "D");
		tFileExcelSheetOutput_1.setDataColumnPosition(4, "E");
		tFileExcelSheetOutput_1.setDataColumnPosition(5, "F");
		tFileExcelSheetOutput_1.setDataColumnPosition(6, "F");
		tFileExcelSheetOutput_1.setDataColumnPosition(7, "A");
		tFileExcelSheetOutput_1.setReuseExistingStylesFromFirstWrittenRow(false);
		tFileExcelSheetOutput_1.setSetupCellStylesForAllColumns(false);
		tFileExcelSheetOutput_1.setReuseFirstRowHeight(false);
		// configure cell formats
		// columnIndex: 1, format: "#,##0", talendType: Long
		tFileExcelSheetOutput_1.setDataFormat(1, "#,##0");
		// columnIndex: 2, format: "@", talendType: String
		tFileExcelSheetOutput_1.setDataFormat(2, "@");
		// columnIndex: 3, format: "dd.mm.yyyy hh:mm", talendType: Date
		tFileExcelSheetOutput_1.setDataFormat(3, "dd.mm.yyyy hh:mm");
		// columnIndex: 5, format: "#,##0.0", talendType: BigDecimal
		tFileExcelSheetOutput_1.setDataFormat(5, "#,##0.0");
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		tFileExcelSheetOutput_1.setWriteZeroDateAsNull(true);
		tFileExcelSheetOutput_1.setForbidWritingInProtectedCells(false);
		// configure auto size columns and group rows by and comments
		// config column 0
		tFileExcelSheetOutput_1.setAutoSizeColumn(0);
		// config column 1
		tFileExcelSheetOutput_1.setAutoSizeColumn(1);
		// config column 2
		tFileExcelSheetOutput_1.setAutoSizeColumn(2);
		// config column 3
		tFileExcelSheetOutput_1.setAutoSizeColumn(3);
		// config column 4
		tFileExcelSheetOutput_1.setAutoSizeColumn(4);
		// config column 5
		tFileExcelSheetOutput_1.setAutoSizeColumn(5);
		// config column 6
		tFileExcelSheetOutput_1.setColumnValueAsComment(6);
		// config column 7
		tFileExcelSheetOutput_1.setColumnValueAsLink(7);
		// fill schema names into the header object array
		Object[] header_tFileExcelSheetOutput_1 = new Object[8];
		header_tFileExcelSheetOutput_1[0] = "My Integer Value";
		header_tFileExcelSheetOutput_1[1] = "What ever I want here";
		header_tFileExcelSheetOutput_1[2] = "string_value";
		header_tFileExcelSheetOutput_1[3] = "date_value";
		header_tFileExcelSheetOutput_1[4] = "bool_value";
		header_tFileExcelSheetOutput_1[5] = "bigdecimal_value";
		header_tFileExcelSheetOutput_1[6] = "comment";
		header_tFileExcelSheetOutput_1[7] = "link";
		// write header
		try {
			tFileExcelSheetOutput_1
					.writeRow(header_tFileExcelSheetOutput_1);
		} catch (Exception e) {
			tFileExcelSheetOutput_1.error(
					"Write header failed: " + e.getMessage(), e);
			globalMap.put("tFileExcelSheetOutput_1_ERROR_MESSAGE",
					"Error in header:" + e.getMessage());
			throw e;
		}
		int nb_line_tFileExcelSheetOutput_1 = 0;
		Object[] dataset_tFileExcelSheetOutput_1 = new Object[8];
		dataset_tFileExcelSheetOutput_1[0] = 12345;
		dataset_tFileExcelSheetOutput_1[1] = 99999999l;
		dataset_tFileExcelSheetOutput_1[2] = "Jan";
		dataset_tFileExcelSheetOutput_1[3] = new Date();
		dataset_tFileExcelSheetOutput_1[4] = true;
		dataset_tFileExcelSheetOutput_1[5] = new BigDecimal("0.12345678");
		dataset_tFileExcelSheetOutput_1[6] = "Das ist Kommentar";
		dataset_tFileExcelSheetOutput_1[7] = "http://jan-lolling.de";

		try {
			tFileExcelSheetOutput_1
					.writeRow(dataset_tFileExcelSheetOutput_1);
			nb_line_tFileExcelSheetOutput_1++;
		} catch (Exception e) {
			tFileExcelSheetOutput_1.error(
					"Write data row in line: "
							+ nb_line_tFileExcelSheetOutput_1
							+ " failed: " + e.getMessage(), e);
			globalMap.put("tFileExcelSheetOutput_1_ERROR_MESSAGE",
					"Write data row in line: "
							+ nb_line_tFileExcelSheetOutput_1
							+ " failed: " + e.getMessage());
			throw e;
		}
		
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_2"));
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_1
					.setOutputFile("/var/testdata/excel/test10/excel_types_out.xlsx");
			tFileExcelWorkbookSave_1.createDirs();
			globalMap.put("tFileExcelWorkbookSave_1_COUNT_SHEETS",
					tFileExcelWorkbookSave_1.getWorkbook()
							.getNumberOfSheets());
			tFileExcelWorkbookSave_1.writeWorkbook();
			// release the memory
			globalMap.put("tFileExcelWorkbookSave_1_FILENAME",
					tFileExcelWorkbookSave_1.getOutputFile());
			globalMap.remove("workbook_tFileExcelWorkbookOpen_2");
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookSave_1_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}
		assertTrue(true);
	}
	
	@Test
	public void testWriteReadFormulas() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen.createEmptyXLSXWorkbook();
		tFileExcelWorkbookOpen.initializeWorkbook();
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName("test_formula");
		tFileExcelSheetOutput_1.resetCache();
		Object[] dataset_tFileExcelSheetOutput_1 = new Object[4];
		dataset_tFileExcelSheetOutput_1[0] = 5;
		dataset_tFileExcelSheetOutput_1[1] = "x";
		dataset_tFileExcelSheetOutput_1[2] = "=A{row}*2";
		dataset_tFileExcelSheetOutput_1[3] = "=CONCATENATE(B{row},A{row})"; // use , instead of ; 
		tFileExcelSheetOutput_1.writeRow(dataset_tFileExcelSheetOutput_1);
		de.jlo.talendcomp.excel.SpreadsheetInput tFileExcelSheetInput_1 = new de.jlo.talendcomp.excel.SpreadsheetInput();
		tFileExcelSheetInput_1.setWorkbook(tFileExcelWorkbookOpen.getWorkbook());
		tFileExcelSheetInput_1.useSheet("test_formula", false);
		int rowindex = 0;
		int a = 0;
		int b = 0;
		String c = null;
		while (tFileExcelSheetInput_1.readNextRow()) {
			a = tFileExcelSheetInput_1.getIntegerCellValue(0, true, false);
			System.out.println("a=" + a);
			b = tFileExcelSheetInput_1.getIntegerCellValue(2, true, false);
			System.out.println("b=" + b);
			c = tFileExcelSheetInput_1.getStringCellValue(3, true, false, false);
			System.out.println("c=" + c);
			rowindex++;
		}
		tFileExcelWorkbookOpen.setOutputFile("/var/testdata/excel/excel_formula_in_out.xlsx");
		tFileExcelWorkbookOpen.writeWorkbook();
		assertEquals(1, rowindex);
		assertEquals("a wrong", 5, a);
		assertEquals("b wrong", 10, b);
		assertEquals("c wrong", "x5", c);
	}
	
	@Test
	public void testAppendDataAndReuseStyles() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen.setInputFile("/Data/Talend/testdata/excel/CellOutput_test/sheet_with_styles.xlsx");
		tFileExcelWorkbookOpen.initializeWorkbook();
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName(0);
		tFileExcelSheetOutput_1.resetCache();
		tFileExcelSheetOutput_1.setRowStartIndex(tFileExcelSheetOutput_1.getSheet().getLastRowNum() + 1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(false);
		// configure cell positions
		tFileExcelSheetOutput_1.setColumnStart(0);
		tFileExcelSheetOutput_1.setAppend(true);
		tFileExcelSheetOutput_1.setReuseExistingStylesFromFirstWrittenRow(true);
		tFileExcelSheetOutput_1.setSetupCellStylesForAllColumns(true);
		tFileExcelSheetOutput_1.setReuseFirstRowHeight(true);
		// configure cell formats
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		tFileExcelSheetOutput_1.setWriteZeroDateAsNull(true);
		tFileExcelSheetOutput_1.setForbidWritingInProtectedCells(false);

		Object[] dataset_tFileExcelSheetOutput_1 = new Object[2];
		for (int i = 0; i < 10; i++) {
			dataset_tFileExcelSheetOutput_1[0] = 100 + i;
			dataset_tFileExcelSheetOutput_1[1] = 15 + i;
			tFileExcelSheetOutput_1.writeRow(dataset_tFileExcelSheetOutput_1);
		}
		
		tFileExcelSheetOutput_1.extendCellRangesForConditionalFormattings();
		
		tFileExcelWorkbookOpen.setOutputFile("/Data/Talend/testdata/excel/CellOutput_test/sheet_with_styles_appended.xlsx");
		tFileExcelWorkbookOpen.writeWorkbook();
	}
	
	@Test
	public void testWriteReadEncrypted() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen.createEmptyXLSXWorkbook();
		tFileExcelWorkbookOpen.initializeWorkbook();
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName("sheet");
		tFileExcelSheetOutput_1.resetCache();
		tFileExcelSheetOutput_1.setAppend(true);
		Object[] dataset_tFileExcelSheetOutput_1 = new Object[2];
		for (int i = 0; i < 10; i++) {
			dataset_tFileExcelSheetOutput_1[0] = 100 + i;
			dataset_tFileExcelSheetOutput_1[1] = 15 + i;
			tFileExcelSheetOutput_1.writeRow(dataset_tFileExcelSheetOutput_1);
		}
		dataset_tFileExcelSheetOutput_1[0] = null;
		dataset_tFileExcelSheetOutput_1[1] = null;
		tFileExcelSheetOutput_1.writeRow(dataset_tFileExcelSheetOutput_1);
		assertEquals(9, tFileExcelSheetOutput_1.detectCurrentSheetLastNoneEmptyRowIndex());
		String targetFile = "/Data/Talend/testdata/excel/test_encrypted/test_encrypted.xlsx";
		tFileExcelWorkbookOpen.setOutputFile(targetFile);
		tFileExcelWorkbookOpen.writeWorkbookEncrypted("secret");
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen2 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen2.setInputFile(targetFile);
		tFileExcelWorkbookOpen2.setPassword("secret");
		tFileExcelWorkbookOpen2.initializeWorkbook();
		assertTrue(true);
	}
	
	@Test
	public void testDetectLastNoneEmptyRow() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen.setInputFile("/var/testdata/excel/test1/cf-result.xlsx");
		tFileExcelWorkbookOpen.initializeWorkbook();
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName(0);
		tFileExcelSheetOutput_1.resetCache();
		tFileExcelSheetOutput_1.setRowStartIndex(2-1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(false);
		int lastRowIndex = tFileExcelSheetOutput_1.detectCurrentSheetLastNoneEmptyRowIndex();
		System.out.println("lastRowIndex = " + lastRowIndex);
		assertEquals(100, lastRowIndex);
	}
	
}
