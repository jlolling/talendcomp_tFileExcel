package de.jlo.talendcomp.excel;

import static org.junit.Assert.assertTrue;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.junit.Test;

public class TestExcelSheetOutput {
	
	@Test
	public void testWriteZeroDate() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		tFileExcelWorkbookOpen_1.setInputFile(
				"/var/testdata/excel/time.xlsx", false);
		tFileExcelWorkbookOpen_1.initializeWorkbook();
		
		final de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setDebug(false);
		tFileExcelSheetOutput_1.setWorkbook(tFileExcelWorkbookOpen_1.getWorkbook());
		tFileExcelSheetOutput_1.setTargetSheetName("out");
		tFileExcelSheetOutput_1.initializeSheet();
		tFileExcelSheetOutput_1.setRowStartIndex(1 - 1);
		tFileExcelSheetOutput_1.setFirstRowIsHeader(false);
		// configure cell positions
		tFileExcelSheetOutput_1.setColumnStart("A");
		tFileExcelSheetOutput_1.setReuseExistingStyles(false);
		tFileExcelSheetOutput_1.setSetupCellStylesForAllColumns(false);
		tFileExcelSheetOutput_1.setReuseFirstRowHeight(false);
		// configure cell formats
		// columnIndex: 0, format: "dd.MM.yyyy", talendType: Date
		tFileExcelSheetOutput_1.setDataFormat(0, "dd.MM.yyyy");
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		tFileExcelSheetOutput_1.setWriteZeroDateAsNull(true);
		String dateStr = "0000-00-00";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		Date zeroDate = sdf.parse(dateStr);
		Object[] row = new Object[2];
		row[0] = zeroDate;
		row[1] = new Date();
		tFileExcelSheetOutput_1.writeRow(row);

		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1.setWorkbook(tFileExcelWorkbookOpen_1.getWorkbook());
		tFileExcelWorkbookSave_1.evaluateAllFormulars();
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

}
