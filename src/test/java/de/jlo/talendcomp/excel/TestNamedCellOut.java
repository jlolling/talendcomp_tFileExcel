package de.jlo.talendcomp.excel;

import java.util.HashMap;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;

public class TestNamedCellOut {
	
	private Map<String, Object> globalMap = new HashMap<>();

	@Before
	public void setUp() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookOpen_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		tFileExcelWorkbookOpen_1.setCreateStreamingXMLWorkbook(false);
		try {
			// read a excel file as template
			// this file file will not used as output file
			tFileExcelWorkbookOpen_1.setInputFile(
					"/var/testdata/excel/test9/template.xlsx", true);
			tFileExcelWorkbookOpen_1.initializeWorkbook();
		} catch (Exception e) {
			tFileExcelWorkbookOpen_1.error(
					"Intialize workbook from file failed: "
							+ e.getMessage(), e);
			globalMap.put("tFileExcelWorkbookOpen_1_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_1",
				tFileExcelWorkbookOpen_1.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_1_COUNT_SHEETS",
				tFileExcelWorkbookOpen_1.getWorkbook()
						.getNumberOfSheets());

	}

	@Test
	public void testWriteNamedCellWithStyle() throws Exception {
		de.jlo.talendcomp.excel.SpreadsheetOutput tFileExcelNamedCellOutput_1 = new de.jlo.talendcomp.excel.SpreadsheetOutput();
		tFileExcelNamedCellOutput_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_1"));
		tFileExcelNamedCellOutput_1
				.setForbidWritingInProtectedCells(false);
		// configure cell positions
		// row counter
		int nb_cells_tFileExcelNamedCellOutput_1 = 0;
		class NullCheck_tFileExcelNamedCellOutput_1 {

			public boolean isNotNull(Object o) {
				return o != null;
			}

		}
		NullCheck_tFileExcelNamedCellOutput_1 nc_tFileExcelNamedCellOutput_1 = new NullCheck_tFileExcelNamedCellOutput_1();
		
		try {
			boolean cellExists = tFileExcelNamedCellOutput_1
					.writeNamedCellValue("p11jhnomb",
							"Jan Lolling");
		} catch (Exception e) {
			tFileExcelNamedCellOutput_1.error(
					"Write flow failed:" + e.getMessage(),
					e);
			globalMap
					.put("tFileExcelNamedCellOutput_1_ERROR_MESSAGE",
							"Write flow failed:"
									+ e.getMessage());
			throw e;
		}

		de.jlo.talendcomp.excel.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.jlo.talendcomp.excel.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_1"));
		tFileExcelWorkbookSave_1.evaluateAllFormulars();
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_1
					.setOutputFile("/var/testdata/excel/test9/named_cell_tests/test_named_cells.xlsx");
			tFileExcelWorkbookSave_1.createDirs();
			globalMap.put("tFileExcelWorkbookSave_1_COUNT_SHEETS",
					tFileExcelWorkbookSave_1.getWorkbook()
							.getNumberOfSheets());
			tFileExcelWorkbookSave_1.writeWorkbook();
			// release the memory
			globalMap.put("tFileExcelWorkbookSave_1_FILENAME",
					tFileExcelWorkbookSave_1.getOutputFile());
			globalMap.remove("workbook_tFileExcelWorkbookOpen_1");
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookSave_1_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

	}

}
