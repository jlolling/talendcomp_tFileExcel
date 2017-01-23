package de.cimt.talendcomp.excel;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

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
	public void testReadStreamingCheckContent() {
		String file = "/Volumes/Data/Talend/testdata/excel/test2/store_report_simple.xlsx";
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
			out.setInputFile("/Volumes/Data/Talend/testdata/excel/copied_cells/Wiser_Pricing_Recommendations_Template.xlsx", true);
			out.initializeWorkbook();
			out.setTargetSheetName("Recommended Actions");
			out.initializeSheet();
			out.setOutputFile("/Volumes/Data/Talend/testdata/excel/copied_cells/Wiser_Pricing_Recommendations_Result.xlsx");
			out.setRowStartIndex(1);
			out.setReuseExistingStyles(true);
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
		tFileExcelSheetInput_2.setNumberFormatLocale("en", true);
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
	
}
