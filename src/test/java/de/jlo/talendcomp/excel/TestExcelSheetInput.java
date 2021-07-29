package de.jlo.talendcomp.excel;

import org.junit.Test;

public class TestExcelSheetInput {
	
	@Test
	public void testReadExcelFormula() throws Exception {
		String file = "/home/jan-lolling/Desktop/OTK-UK_20210729_0938_Pricing Data V4 August 28.07.2021.xlsx";
		SpreadsheetFile ssf = new SpreadsheetFile();
		ssf.setInputFile(file);
		ssf.initializeWorkbook();
		SpreadsheetInput tFileExcelSheetInput_1 = new de.jlo.talendcomp.excel.SpreadsheetInput();
		tFileExcelSheetInput_1.setWorkbook(ssf.getWorkbook());
		tFileExcelSheetInput_1.useSheet(0, true);
		tFileExcelSheetInput_1.setStopAtMissingRow(true);
		tFileExcelSheetInput_1.setRowStartIndex(2 - 1);
		// configure cell positions
		tFileExcelSheetInput_1.setFormatLocale("en", true);
		tFileExcelSheetInput_1.setParseDateFromVisibleString(false);
		tFileExcelSheetInput_1.setLenientDateParsing(true);
		tFileExcelSheetInput_1.setReturnZeroDateAsNull(true);
		tFileExcelSheetInput_1.setReturnURLInsteadOfName(false);
		tFileExcelSheetInput_1.setConcatenateLabelUrl(false);
		if (tFileExcelSheetInput_1.readNextRow()) {
			String value = tFileExcelSheetInput_1.getStringCellValue(8, false, false, false);
			System.out.println(value);
		}
	}

}
