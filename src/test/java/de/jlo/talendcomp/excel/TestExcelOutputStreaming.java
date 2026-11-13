package de.jlo.talendcomp.excel;

import java.util.HashMap;
import java.util.Map;

public class TestExcelOutputStreaming {
	
	public static Map<String, Object> globalMap = new HashMap<>();
	
	public static class row2Struct {

		public String newColumn;

		public String getNewColumn() {
			return this.newColumn;
		}

		public String newColumn1;

		public String getNewColumn1() {
			return this.newColumn1;
		}

		public String newColumn2;

		public String getNewColumn2() {
			return this.newColumn2;
		}

		public String newColumn3;

		public String getNewColumn3() {
			return this.newColumn3;
		}

	}
	
	
	public void testAppend() throws Exception {
		
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_2 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_2.setZipBombWarningThreshold(0.005d);
		tFileExcelWorkbookOpen_2.setCreateStreamingXMLWorkbook(true);
		try {
			// read a excel file as template
			// this file file will not used as output file
			tFileExcelWorkbookOpen_2.setInputFile("/var/testdata/excel/test_streaming.xlsx",
					true);
			tFileExcelWorkbookOpen_2.initializeWorkbook();
		} catch (Exception e) {
			tFileExcelWorkbookOpen_2.error("Intialize workbook from file failed: " + e.getMessage(), e);
			globalMap.put("tFileExcelWorkbookOpen_2_ERROR_MESSAGE", e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_2", tFileExcelWorkbookOpen_2.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_2_COUNT_SHEETS",
				tFileExcelWorkbookOpen_2.getWorkbook().getNumberOfSheets());

		
		final de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_2 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_2.setWorkbook(
				(org.apache.poi.ss.usermodel.Workbook) globalMap.get("workbook_tFileExcelWorkbookOpen_2"));
		tFileExcelSheetOutput_2.setTargetSheetName("sheet");
		globalMap.put("tFileExcelSheetOutput_2_SHEET_NAME", tFileExcelSheetOutput_2.getTargetSheetName());
		tFileExcelSheetOutput_2.resetCache();
		int startRowIndex_tFileExcelSheetOutput_2 = 5 - 1;
		int currentSheetLastRowIndex_tFileExcelSheetOutput_2 = tFileExcelSheetOutput_2
				.detectCurrentSheetLastNoneEmptyRowIndex() + 1;
		tFileExcelSheetOutput_2.setRowStartIndex(
				startRowIndex_tFileExcelSheetOutput_2 > currentSheetLastRowIndex_tFileExcelSheetOutput_2
						? startRowIndex_tFileExcelSheetOutput_2
						: currentSheetLastRowIndex_tFileExcelSheetOutput_2);
		tFileExcelSheetOutput_2.setTemplateRowIndexForStyles(startRowIndex_tFileExcelSheetOutput_2);
		tFileExcelSheetOutput_2.setFirstRowIsHeader(false);
		// configure cell positions
		tFileExcelSheetOutput_2.setColumnStart("C");
		tFileExcelSheetOutput_2.setAppend(true);
		tFileExcelSheetOutput_2.setReuseExistingStylesFromFirstWrittenRow(false);
		tFileExcelSheetOutput_2.setSetupCellStylesForAllColumns(false);
		tFileExcelSheetOutput_2.setReuseFirstRowHeight(false);
		// configure cell formats
		// columnIndex: 0, name: newColumn, format: , talendType: String
		// columnIndex: 1, name: newColumn1, format: , talendType: String
		// columnIndex: 2, name: newColumn2, format: , talendType: String
		// columnIndex: 3, name: newColumn3, format: , talendType: String
		tFileExcelSheetOutput_2.setWriteNullValues(false);
		tFileExcelSheetOutput_2.setWriteZeroDateAsNull(true);
		tFileExcelSheetOutput_2.setForbidWritingInProtectedCells(false);
		// configure auto size columns and group rows by and comments
		// config column 0
		// config column 1
		// config column 2
		// config column 3
		// row counter
		int nb_line_tFileExcelSheetOutput_2 = 0;
		
		
		row2Struct row2 = new row2Struct();
		// fill schema data into the object array
		Object[] dataset_tFileExcelSheetOutput_2 = new Object[4];
		dataset_tFileExcelSheetOutput_2[0] = row2.newColumn;
		dataset_tFileExcelSheetOutput_2[1] = row2.newColumn1;
		dataset_tFileExcelSheetOutput_2[2] = row2.newColumn2;
		dataset_tFileExcelSheetOutput_2[3] = row2.newColumn3;
		// write dataset
		try {
			tFileExcelSheetOutput_2.writeRow(dataset_tFileExcelSheetOutput_2);
			nb_line_tFileExcelSheetOutput_2++;
		} catch (Exception e) {
			tFileExcelSheetOutput_2.error("Write data row in line: " + nb_line_tFileExcelSheetOutput_2
					+ " failed: " + e.getMessage(), e);
			globalMap.put("tFileExcelSheetOutput_2_ERROR_MESSAGE", "Write data row in line: "
					+ nb_line_tFileExcelSheetOutput_2 + " failed: " + e.getMessage());
			throw e;
		}

		tFileExcelSheetOutput_2.setupColumnSize();
		tFileExcelSheetOutput_2.closeLastGroup();
		globalMap.put("tFileExcelSheetOutput_2_NB_LINE", nb_line_tFileExcelSheetOutput_2);
		globalMap.put("tFileExcelSheetOutput_2_LAST_ROW_INDEX",
				tFileExcelSheetOutput_2.detectCurrentSheetLastNoneEmptyRowIndex() + 1);

		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_2 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_2.setWorkbook(
				(org.apache.poi.ss.usermodel.Workbook) globalMap.get("workbook_tFileExcelWorkbookOpen_2"));
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_2.setOutputFile("/var/testdata/excel/test_streaming2");
			tFileExcelWorkbookSave_2.createDirs();
			globalMap.put("tFileExcelWorkbookSave_2_COUNT_SHEETS",
					tFileExcelWorkbookSave_2.getWorkbook().getNumberOfSheets());
			tFileExcelWorkbookSave_2.writeWorkbook();
			// release the memory
			globalMap.put("tFileExcelWorkbookSave_2_FILENAME", tFileExcelWorkbookSave_2.getOutputFile());
			globalMap.remove("workbook_tFileExcelWorkbookOpen_2");
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookSave_2_ERROR_MESSAGE", e.getMessage());
			throw e;
		}

	}

}
