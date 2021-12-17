package de.jlo.talendcomp.excel;

import java.util.HashMap;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;

public class TestReferencedCellOutput {
	
	private Map<String, Object> globalMap = new HashMap<>();
	
	@Before
	public void createWorkbook() throws Exception {
		// create workbook
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setZipBombWarningThreshold(0.005d);
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		try {
			// create empty XLSX workbook
			tFileExcelWorkbookOpen_1.createEmptyXLSXWorkbook();
			tFileExcelWorkbookOpen_1.initializeWorkbook();
		} catch (Exception e) {
			tFileExcelWorkbookOpen_1.error("Intialize empty workbook failed: " + e.getMessage(), e);
			globalMap.put("tFileExcelWorkbookOpen_1_ERROR_MESSAGE", e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_1", tFileExcelWorkbookOpen_1.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_1_COUNT_SHEETS",
				tFileExcelWorkbookOpen_1.getWorkbook().getNumberOfSheets());
		// create sheet
		final de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelSheetOutput_1.setWorkbook(
				(org.apache.poi.ss.usermodel.Workbook) globalMap.get("workbook_tFileExcelWorkbookOpen_1"));
		tFileExcelSheetOutput_1.setTargetSheetName("Test");
		globalMap.put("tFileExcelSheetOutput_1_SHEET_NAME", tFileExcelSheetOutput_1.getTargetSheetName());
	}
	
	@Test
	public void testWriteReferencedCell() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelReferencedCellOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelReferencedCellOutput_1.setWorkbook(
				(org.apache.poi.ss.usermodel.Workbook) globalMap.get("workbook_tFileExcelWorkbookOpen_1"));
		tFileExcelReferencedCellOutput_1.setForbidWritingInProtectedCells(false);
		try {
			String cellRef = null; //"A1";
			int sheetIndex = 0;
			int row = 1;
			int col = 0;
			String value = "test";
			String comment = "comment";
			String autor = "autor";
			// address target cell
			tFileExcelReferencedCellOutput_1.setCommentAuthor(autor);
			tFileExcelReferencedCellOutput_1.writeReferencedCellValue(
					cellRef,
					sheetIndex, 
					row, 
					col, 
					value,
					comment, 
					false);
		} catch (Exception e) {
			tFileExcelReferencedCellOutput_1.error("Write cell failed:" + e.getMessage(), e);
			globalMap.put("tFileExcelReferencedCellOutput_1_ERROR_MESSAGE",
					"Write cell failed:" + e.getMessage());
			throw e;
		}
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1.setWorkbook(
				(org.apache.poi.ss.usermodel.Workbook) globalMap.get("workbook_tFileExcelWorkbookOpen_1"));
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_1.setOutputFile("/var/testdata/excel/write_referenced_cells/test2-result");
			tFileExcelWorkbookSave_1.createDirs();
			globalMap.put("tFileExcelWorkbookSave_1_COUNT_SHEETS",
					tFileExcelWorkbookSave_1.getWorkbook().getNumberOfSheets());
			tFileExcelWorkbookSave_1.writeWorkbook();
			// release the memory
			globalMap.put("tFileExcelWorkbookSave_1_FILENAME", tFileExcelWorkbookSave_1.getOutputFile());
			globalMap.remove("workbook_tFileExcelWorkbookOpen_1");
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookSave_1_ERROR_MESSAGE", e.getMessage());
			throw e;
		}
		
	}

}
