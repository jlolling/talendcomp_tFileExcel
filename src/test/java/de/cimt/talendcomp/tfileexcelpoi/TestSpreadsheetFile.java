package de.cimt.talendcomp.tfileexcelpoi;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

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
		
}
