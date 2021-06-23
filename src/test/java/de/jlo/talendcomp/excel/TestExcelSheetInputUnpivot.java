package de.jlo.talendcomp.excel;

import static org.junit.Assert.assertEquals;

import org.junit.Before;
import org.junit.Test;

public class TestExcelSheetInputUnpivot {

	SpreadsheetInput tFileExcelSheetInput_2 = null;
	
	@Before
	public void testReadXls() throws Exception {
		String path = TestUtil.writeResourceToFile("/test_unpivot.xlsx", "/tmp/");
		System.out.println("Use test file: " + path);
		System.out.println(System.getProperty("java.version"));
		tFileExcelSheetInput_2 = new SpreadsheetInput();
		tFileExcelSheetInput_2.setInputFile(path);
		tFileExcelSheetInput_2.initializeWorkbook();
		tFileExcelSheetInput_2.useSheet(0, false);
		tFileExcelSheetInput_2.setStopAtMissingRow(true);
		// configure cell positions
		tFileExcelSheetInput_2.setDataColumnPosition(0, "A");
		tFileExcelSheetInput_2.setFormatLocale("en", true);
		tFileExcelSheetInput_2.setDefaultDateFormat("yyyy-MM-dd HH:mm:ss");
		tFileExcelSheetInput_2.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_2.setConcatenateLabelUrl(false);
	}

	@Test
	public void testUnpivot1() throws Exception {
		tFileExcelSheetInput_2.setRowStartIndex(2);
		SpreadsheetInputUnpivot up = new SpreadsheetInputUnpivot();
		up.setHeaderRowIndex(2-1);
		up.setUnpivotColumnRangeStartIndex(0);
		int count = 0;
		while (tFileExcelSheetInput_2.readNextRow()) {
			String agStr = tFileExcelSheetInput_2.getStringCellValue(0, true, true, false);
			System.out.println(agStr);
			up.checkAndInitialize(tFileExcelSheetInput_2);
			up.normalizeValuesOfCurrentRow();
			while (up.nextNormalizedRow()) {
				System.out.print("row-index: " + up.getCurrentOriginalRowIndex());
				System.out.print(" column-index: " + up.getCurrentOriginalColumnIndex());
				System.out.print(" header: " + up.getCurrentHeaderAsString());
				System.out.println(" value: " + up.getCurrentValueAsLong());
				count++;
			}
		}
		assertEquals(12, count);
	}
	
	@Test
	public void testUnpivot2() throws Exception {
		tFileExcelSheetInput_2.setRowStartIndex(9);
		SpreadsheetInputUnpivot up = new SpreadsheetInputUnpivot();
		up.setHeaderRowIndex(9-1);
		up.setUnpivotColumnRangeStartIndex("D");
		up.setUnpivotColumnRangeEndIndex("E");
		int count = 0;
		while (tFileExcelSheetInput_2.readNextRow()) {
			String agStr = tFileExcelSheetInput_2.getStringCellValue(0, true, true, false);
			System.out.println(agStr);
			up.checkAndInitialize(tFileExcelSheetInput_2);
			up.normalizeValuesOfCurrentRow();
			while (up.nextNormalizedRow()) {
				System.out.print("row-index: " + up.getCurrentOriginalRowIndex());
				System.out.print(" column-index: " + up.getCurrentOriginalColumnIndex());
				System.out.print(" header: " + up.getCurrentHeaderAsString());
				System.out.println(" value: " + up.getCurrentValueAsLong());
				count++;
			}
		}
		assertEquals(6, count);
	}
	
}
