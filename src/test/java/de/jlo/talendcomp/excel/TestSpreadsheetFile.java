package de.jlo.talendcomp.excel;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;


public class TestSpreadsheetFile {
	
	protected Workbook workbook;

	@Before
	public void setUp() throws Exception {
		
	}
	
	@Test
	public void testReadXls() throws Exception {
		SpreadsheetInput tFileExcelSheetInput_2 = new SpreadsheetInput();
		tFileExcelSheetInput_2.setInputFile("/Data/Talend/testdata/excel/4803e947fa70cfcd828079fe857f8a1b.xls");
		tFileExcelSheetInput_2.initializeWorkbook();
		tFileExcelSheetInput_2.useSheet(0);
		tFileExcelSheetInput_2.setStopAtMissingRow(false);
		tFileExcelSheetInput_2.setRowStartIndex(3);
		// configure cell positions
		tFileExcelSheetInput_2.setDataColumnPosition(0, "A");
		tFileExcelSheetInput_2.setFormatLocale("en", true);
		tFileExcelSheetInput_2.setDefaultDateFormat("yyyy-MM-dd HH:mm:ss");
		tFileExcelSheetInput_2.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_2.setConcatenateLabelUrl(false);
		while (tFileExcelSheetInput_2.readNextRow()) {
			String agStr = tFileExcelSheetInput_2.getStringCellValue(0, true, true, false);
			System.out.println(agStr);
		}
	}

	@Test
	public void testReadStreamingCheckContent() {
		String file = "/var/testdata/excel/test2/store_report_simple.xlsx";
		SpreadsheetFile sf = new SpreadsheetFile();
		sf.setCreateStreamingXMLWorkbook(true);
		try {
			sf.setInputFile(file,true);
			sf.initializeWorkbook();
		} catch (Exception e) {
			fail("Read file failed:" + e.getMessage());
		}
		SpreadsheetInput si = new SpreadsheetInput();
		si.setWorkbook(sf.getWorkbook());
		try {
			si.useSheet(0);
		} catch (Exception e) {
			fail("use sheet failed: " + e.getMessage());
		}
		Row row = si.getRow(0);
		if (row == null) {
			fail("Row 1 does not exists");
		}
		int rowNum = 0;
		while (si.readNextRow()) {
			rowNum++;
		}
		assertTrue("Not correct row num:" + rowNum, rowNum == 972001);
	}
	
	@Test
	public void testReadStreamingLastRowNum() {
		String file = "/Volumes/Data/Talend/testdata/excel/test2/store_report.xlsx";
		SpreadsheetFile sf = new SpreadsheetFile();
		sf.setCreateStreamingXMLWorkbook(true);
		try {
			sf.setInputFile(file, true);
			sf.initializeWorkbook();
			workbook = sf.getWorkbook();
		} catch (Exception e) {
			fail("Read file failed:" + e.getMessage());
		}
		SpreadsheetInput si = new SpreadsheetInput();
		si.setWorkbook(workbook);
		try {
			si.useSheet("Store 11");
		} catch (Exception e) {
			fail("use sheet failed: " + e.getMessage());
		}
		assertTrue("No rows found", si.getLastRowNum() > 0);
	}
	
	@Test
	public void testInsertRow() {
		SpreadsheetOutput out = new SpreadsheetOutput();
		try {
			out.createEmptyXLSXWorkbook();
			out.initializeWorkbook();
			out.initializeSheet();
			out.setOutputFile("/var/testdata/excel/excel_shift_test.xlsx");
			out.freezeAt(0, 1);
			for (int r = 0; r < 9; r++) {
				Object[] row = new Object[20];
				for (int c = 0; c < 2; c++) {
					if (c == 0) {
						row[c] = 100 + r;
					} else {
						row[c] = "=10+A{row}";
					}
				}
				out.writeRow(row);
			}
			// now insert rows
			out.setRowStartIndex(2);
			for (int r = 2; r < 4; r++) {
				Object[] row = new Object[20];
				for (int c = 0; c < 2; c++) {
					if (c == 0) {
						row[c] = 400 + r;
					} else {
						row[c] = "=10+A{row}";
					}
				}
				out.shiftCurrentRow();
				out.writeRow(row);
			}
			out.writeWorkbook();
		} catch (Exception e) {
			e.printStackTrace();
			fail("Initialise workbook failed: " + e.getMessage());
		}
	}

	@Test
	public void testWithAppendDataValidations() {
		SpreadsheetOutput out = new SpreadsheetOutput();
		out.setDebug(true);
		out.setSetupCellStylesForAllColumns(true);
		try {
			out.setInputFile("/Data/Talend/testdata/excel/copied_cells/Wiser_Pricing_Recommendations_Template.xlsx", true);
			out.initializeWorkbook();
			out.setTargetSheetName("Recommended Actions");
			out.initializeSheet();
			out.setOutputFile("/Data/Talend/testdata/excel/copied_cells/Wiser_Pricing_Recommendations_Result.xlsx");
			out.setRowStartIndex(1);
			out.setReuseExistingStylesFromFirstWrittenRow(true);
			for (int r = 0; r < 9; r++) {
				Object[] row = new Object[2];
				for (int c = 0; c < 2; c++) {
					if (c == 0) {
						row[c] = 100 + r;
					} else {
						row[c] = "Produkt-" + r;
					}
				}
				out.writeRow(row);
			}
			out.createDataValidationsForAppendedRows();
			out.writeWorkbook();
		} catch (Exception e) {
			e.printStackTrace();
			fail("Initialise workbook failed: " + e.getMessage());
		}
	}

	@Test
	public void testReadDuration() throws Exception {
		SpreadsheetInput tFileExcelSheetInput_2 = new SpreadsheetInput();
		tFileExcelSheetInput_2.setInputFile("/var/testdata/excel/time.xls");
		tFileExcelSheetInput_2.initializeWorkbook();
		tFileExcelSheetInput_2.useSheet(0);
		tFileExcelSheetInput_2.setStopAtMissingRow(false);
		tFileExcelSheetInput_2.setRowStartIndex(0);
		// configure cell positions
		tFileExcelSheetInput_2.setDataColumnPosition(0, "B");
		tFileExcelSheetInput_2.setDataColumnPosition(1, "B");
		tFileExcelSheetInput_2.setFormatLocale("en", true);
		tFileExcelSheetInput_2.setDefaultDateFormat("yyyy-MM-dd HH:mm:ss");
		tFileExcelSheetInput_2.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_2.setConcatenateLabelUrl(false);
		while (tFileExcelSheetInput_2.readNextRow()) {
			Date durationDate = tFileExcelSheetInput_2.getDurationCellValue(0, true, false, "HH:mm:ss");
			String agStr = tFileExcelSheetInput_2.getStringCellValue(1, true, true, false);
			CellStyle style1 = tFileExcelSheetInput_2.getCellStyle(1);
			String pattern1 = "--";
			if (style1 != null) {
				pattern1 = style1.getDataFormatString();
			}
			System.out.println();
			System.out.println("row: " + tFileExcelSheetInput_2.getCurrentRowIndex() + " value: " + durationDate.getTime() + " pattern: " + pattern1);
			System.out.println("row: " + tFileExcelSheetInput_2.getCurrentRowIndex() + " value: " + agStr + " pattern: " + pattern1);
		}
		assertTrue(true);
	}
	
	@Test
	public void testReadDateLocalised() throws Exception {
		SpreadsheetInput tFileExcelSheetInput_2 = new SpreadsheetInput();
		tFileExcelSheetInput_2.setInputFile("/var/testdata/excel/time.xls");
		tFileExcelSheetInput_2.initializeWorkbook();
		tFileExcelSheetInput_2.useSheet(0);
		tFileExcelSheetInput_2.setStopAtMissingRow(false);
		tFileExcelSheetInput_2.setRowStartIndex(0);
		// configure cell positions
		tFileExcelSheetInput_2.setDataColumnPosition(0, "A");
		tFileExcelSheetInput_2.setFormatLocale("de", true);
		tFileExcelSheetInput_2.setDefaultDateFormat("yyyy-MM-dd HH:mm:ss");
		tFileExcelSheetInput_2.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_2.setConcatenateLabelUrl(false);
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		while (tFileExcelSheetInput_2.readNextRow()) {
			try {
				Date date = tFileExcelSheetInput_2.getDateCellValue(0, true, false, null);
				System.out.println("row: " + (tFileExcelSheetInput_2.getCurrentRowIndex() + 1) + " value: " + sdf.format(date));
			} catch (Exception e) {
				System.out.println("Error in row: " + (tFileExcelSheetInput_2.getCurrentRowIndex() + 1) + " message: " + e.getMessage());
			}
		}
		assertTrue(true);
	}

	@Test
	public void testReadNumberFormatted() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		try {
			// read a excel file as template
			// this file file will not used as output file
			tFileExcelWorkbookOpen_1
					.setInputFile(
							"/Volumes/Data/Talend/testdata/excel/test_double.xls",
							true);
			tFileExcelWorkbookOpen_1.initializeWorkbook();
		} catch (Exception e) {
			tFileExcelWorkbookOpen_1.error(
					"Intialize workbook from file failed: "
							+ e.getMessage(), e);
			throw e;
		}
		
		final de.jlo.talendcomp.excel.SpreadsheetInput tFileExcelSheetInput_1 = new de.jlo.talendcomp.excel.SpreadsheetInput();
		tFileExcelSheetInput_1
				.setWorkbook(tFileExcelWorkbookOpen_1.getWorkbook());
		tFileExcelSheetInput_1.useSheet(0);
		tFileExcelSheetInput_1.setStopAtMissingRow(false);
		tFileExcelSheetInput_1.setRowStartIndex(1 - 1);
		// configure cell positions
		tFileExcelSheetInput_1.setDataColumnPosition(0, "A");
		tFileExcelSheetInput_1.setDataColumnPosition(1, "B");
		tFileExcelSheetInput_1.setDataColumnPosition(2, "A");
		tFileExcelSheetInput_1.setDataColumnPosition(3, "B");
		tFileExcelSheetInput_1.setFormatLocale("fr", false);
		tFileExcelSheetInput_1.setDefaultDateFormat("yyyyMMddHHmmss");
		tFileExcelSheetInput_1.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_1.setConcatenateLabelUrl(false);
		tFileExcelSheetInput_1.setNumberPrecision(0, 5);
		tFileExcelSheetInput_1.setNumberPrecision(1, 5);

		// row counter
		int nb_line_tFileExcelSheetInput_1 = 0;
		while (tFileExcelSheetInput_1.readNextRow()) {
			nb_line_tFileExcelSheetInput_1++;
			try {
				String col1_str = tFileExcelSheetInput_1.getStringCellValue(0,
						true, false, false);
				System.out.println(col1_str);
			} catch (Exception e) {
				throw new Exception(
						"Read column col1_str in row number="
								+ nb_line_tFileExcelSheetInput_1 + " failed:"
								+ e.getMessage(), e);
			}
			try {
				String col2_str = tFileExcelSheetInput_1.getStringCellValue(1,
						true, false, false);
				System.out.println(col2_str);
			} catch (Exception e) {
				throw new Exception(
						"Read column col2_str in row number="
								+ nb_line_tFileExcelSheetInput_1 + " failed:"
								+ e.getMessage(), e);
			}
			try {
				Double col1_double = tFileExcelSheetInput_1.getDoubleCellValue(2,
						true, false);
			} catch (Exception e) {
				tFileExcelSheetInput_1
						.warn("Read column col1_double in row number="
								+ nb_line_tFileExcelSheetInput_1
								+ " error ignored:"
								+ e.getMessage());
			}
			try {
				Double col2_double = tFileExcelSheetInput_1.getDoubleCellValue(3,
						true, false);
			} catch (Exception e) {
				tFileExcelSheetInput_1
						.warn("Read column col2_double in row number="
								+ nb_line_tFileExcelSheetInput_1
								+ " error ignored:"
								+ e.getMessage());
			}
			assertTrue(true);
		}
		
	}
	
}
