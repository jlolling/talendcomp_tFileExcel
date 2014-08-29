
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import de.cimt.talendcomp.tfileexcelpoi.SpreadsheetInput;
import de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput;


public class TextSheetOutput {
	
	private static Map<String, Object> globalMap = new HashMap<String, Object>();

	/**
	 * @param args
	 * @throws ParseException 
	 */
	public static void main(String[] args) throws Exception {
		testTables();
		//testTypes();
	}
	
	public static void testOutputStyled() {
		SpreadsheetOutput out = new SpreadsheetOutput();
		try {
			out.createEmptyXLSWorkbook();
			out.initializeWorkbook();
			out.initializeSheet();
			out.addStyle("odd", "Arial", "10", "", "8", "49", "left", false);
			out.addStyle("even", "Arial", "10", "", "9", "12", "left", false);
			out.setOddRowStyleName("odd");
			out.setEvenRowStyleName("even");
			out.setOutputFile("/Users/jan/test/excel/styled_excel.xls");
			for (int r = 0; r < 9; r++) {
				Object[] row = new Object[1];
				for (int c = 0; c < 1; c++) {
					row[c] = "value:" + r + "-" + c;
				}
				out.writeRow(row);
			}
			out.writeWorkbook();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void printHSSFColors() {
		Map<Integer, HSSFColor> map = HSSFColor.getIndexHash();
		TreeSet<String> set = new TreeSet<String>();
		for (Map.Entry<Integer, HSSFColor> entry : map.entrySet()) {
			int index = entry.getKey();
			String className = entry.getValue().getClass().getName();
			int pos = className.indexOf("$");
			String colorName = className.substring(pos + 1);
			String item = colorName + " index="+index+" class:" + className;
			set.add(item);
		}
		for (String s : set) {
			System.out.println(s);
		}
	}
	
	public static void testMacroIfSurvives() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		try {
			e.setInputFile("/home/jlolling/test/excel/macro_test.xlsm");
			e.initializeWorkbook();
			e.useSheet(1);
			e.setRowStartIndex(5);
			e.setDataColumnPosition(0, "C");
			e.setDataColumnPosition(1, "D");
			e.setDataColumnPosition(2, "E");
			e.setDataColumnPosition(3, "F");
			e.setDataColumnPosition(4, "G");
			e.setDataColumnPosition(5, "H");
			e.setDataColumnPosition(6, "I");
			e.setDataColumnPosition(7, "J");
			int rowIndex = 0;
			while (e.readNextRow()) {
				try {
					Object v0 = e.getIntegerCellValue(0, true, false);
					System.out.print(v0);
					System.out.print("|");
				} catch (Exception e1) {
					System.err.println("rowIndex=" + rowIndex + " column 0");
					e1.printStackTrace();
				}
				
//				Object v1 = e.getStringCellValue(1, true, false, false);
//				System.out.print(v1);
//				System.out.print("|");
//				Object v2 = e.getIntegerCellValue(2, true, false);
//				System.out.print(v2);
//				System.out.print("|");
//				Object v3 = e.getIntegerCellValue(3, true, false);
//				System.out.print(v3);
//				System.out.println("|");
				Object v4 = e.getDateCellValue(4, true, false, "dd.MM.yyyy");
				System.out.print(v4);
				System.out.print("|");
				Object v5 = e.getDoubleCellValue(5, true, false);
				System.out.print(v5);
				System.out.print("|");
				Object v6 = e.getDoubleCellValue(6, true, false);
				System.out.print(v6);
				System.out.print("|");
				Object v7 = e.getStringCellValue(7, true, false, false);
				System.out.print(v7);
				System.out.print("|");
				Object v8 = e.getStringCellValue(8, true, false, false);
				System.out.println(v8);
				System.out.println("-------");
				rowIndex++;
			}
			e.setOutputFile("/home/jlolling/test/excel/macro_test_result.xlsm");
			e.writeWorkbook();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public static void testMassiveOutput() throws Exception {
		SpreadsheetOutput out = new SpreadsheetOutput();
		out.setCreateStreamingXMLWorkbook(true);
		out.createEmptyXLSXWorkbook();
		out.initializeWorkbook();
		out.initializeSheet();
		out.freezeAt(10, 10);
		out.setColumnStart("C");
		out.groupRowsByColumn(2);
		//out.groupRowsByColumn(4);
		out.addColumnGroup("D-G, AA - AG");
		out.setRowStartIndex(5);
		for (int r = 0; r < 50; r++) {
			Object[] row = new Object[100];
			for (int c = 0; c < 100; c++) {
				if (c == 2) {
					row[c] = "group1:" + (r / 5) + "-" + c;
				} else if (c == 4) {
					row[c] = "group2:" + (r / 7) + "-" + c;
				} else {
					row[c] = "!value:" + r + "-" + c;
				}
			}
			out.writeRow(row);
		}
		out.closeLastGroup();
		out.setOutputFile("/var/testdata/excel/large.xlsx");
		out.writeWorkbook();
	}
	
	public static void testEmpty() throws Exception {
		SpreadsheetOutput out = new SpreadsheetOutput();
		out.createEmptyXLSXWorkbook();
		out.initializeWorkbook();
		out.initializeSheet();
		System.out.println("is empty:" + out.isEmpty());
	}

	public static void testXLSCreateComment() throws Exception {
		SpreadsheetOutput e = new SpreadsheetOutput();
		e.createEmptyXLSXWorkbook();
		e.initializeWorkbook();
		e.setTargetSheetName("with_comments");
		e.initializeSheet();
		e.setCommentAuthor("Jan Lolling");
		e.setCommentHeight(4);
		e.setCommentWidth(5);
		e.writeReferencedCellValue(0, 1, "Jan", "Kommentar", null);
		e.writeReferencedCellValue(1, 2, "Feb", "Kommentar", null);
		e.writeReferencedCellValue(2, 0, 2, null, null);
		e.writeReferencedCellValue(2, 1, 5, "toller Wert", null);
		e.writeReferencedCellValue(2, 2, "=A{row}+B3", "Ergebnis", null);
		e.setOutputFile("/private/var/testdata/excel/comments.xlsx");
		e.writeWorkbook();
	}

	public static void testNamedCells() throws Exception {
		SpreadsheetOutput e = new SpreadsheetOutput();
		e.setInputFile("/private/var/testdata/excel/named_cell_tests/template.xlsx");
		e.initializeWorkbook();
		String name = "p11telefo";
		if (e.writeNamedCellValue(name, 1234) == false) {
			throw new Exception("cell " + name + " cannot found");
		}
		e.setOutputFile("/private/var/testdata/excel/named_cell_tests/named_cells_written.xlsx");
		e.writeWorkbook();
	}

	public static void testTypes() throws Exception {

		de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile tFileExcelWorkbookOpen_2 = new de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile();
		tFileExcelWorkbookOpen_2.setCreateStreamingXMLWorkbook(false);
		try {
			// create empty XLSX workbook
			tFileExcelWorkbookOpen_2.createEmptyXLSXWorkbook();
			tFileExcelWorkbookOpen_2.initializeWorkbook();
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookOpen_2_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_2",
				tFileExcelWorkbookOpen_2.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_2_COUNT_SHEETS",
				tFileExcelWorkbookOpen_2.getWorkbook()
						.getNumberOfSheets());

		int nb_line_tFileExcelSheetOutput_1 = 0; 
		de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput();
		tFileExcelSheetOutput_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_2"));
		tFileExcelSheetOutput_1.setTargetSheetName("test_out");
		tFileExcelSheetOutput_1.initializeSheet();
		tFileExcelSheetOutput_1.setRowStartIndex(1 - 1);
		// configure cell positions
		tFileExcelSheetOutput_1.setColumnStart("A");
		// configure cell formats
		// columnIndex: 1, format: "#,##0", talendType: Long
		tFileExcelSheetOutput_1.setDataFormat(1, "#,##0");
		// columnIndex: 3, format: "dd.mm.yyyy hh:mm", talendType: Date
		tFileExcelSheetOutput_1.setDataFormat(3, "dd.mm.yyyy hh:mm");
		// columnIndex: 5, format: "#,##0.0", talendType: BigDecimal
		tFileExcelSheetOutput_1.setDataFormat(5, "#,##0.0");
		tFileExcelSheetOutput_1.setWriteNullValues(false);
		// configure auto size columns
		tFileExcelSheetOutput_1.setAutoSizeColumn(0);
		tFileExcelSheetOutput_1.setAutoSizeColumn(1);
		tFileExcelSheetOutput_1.setAutoSizeColumn(2);
		tFileExcelSheetOutput_1.setAutoSizeColumn(3);
		tFileExcelSheetOutput_1.setAutoSizeColumn(4);
		tFileExcelSheetOutput_1.setAutoSizeColumn(5);
		// fill schema names into the header object array
		Object[] header_tFileExcelSheetOutput_1 = new Object[6];
		header_tFileExcelSheetOutput_1[0] = "int_value";
		header_tFileExcelSheetOutput_1[1] = "long_value";
		header_tFileExcelSheetOutput_1[2] = "string_value";
		header_tFileExcelSheetOutput_1[3] = "date_value";
		header_tFileExcelSheetOutput_1[4] = "bool_value";
		header_tFileExcelSheetOutput_1[5] = "bigdecimal_value";
		// write header
		try {
			tFileExcelSheetOutput_1
					.writeRow(header_tFileExcelSheetOutput_1);
		} catch (Exception e) {
			globalMap.put("tFileExcelSheetOutput_1_ERROR_MESSAGE",
					"Error in header:" + e.getMessage());
			throw e;
		}

		int int_value = 2222;

		long long_value = 10000001;

		String string_value = "Testüöäß\"";

		Date date_value = new Date();

		boolean bool_value = true;

		BigDecimal bigdecimal_value = new BigDecimal("1.23456");

		// fill schema data into the object array
		Object[] dataset_tFileExcelSheetOutput_1 = new Object[6];
		dataset_tFileExcelSheetOutput_1[0] = int_value;
		dataset_tFileExcelSheetOutput_1[1] = long_value;
		dataset_tFileExcelSheetOutput_1[2] = string_value;
		dataset_tFileExcelSheetOutput_1[3] = date_value;
		dataset_tFileExcelSheetOutput_1[4] = bool_value;
		dataset_tFileExcelSheetOutput_1[5] = bigdecimal_value;
		// write dataset
		try {
			tFileExcelSheetOutput_1
					.writeRow(dataset_tFileExcelSheetOutput_1);
			nb_line_tFileExcelSheetOutput_1++;
		} catch (Exception e) {
			globalMap.put("tFileExcelSheetOutput_1_ERROR_MESSAGE",
					"Error in line "
							+ nb_line_tFileExcelSheetOutput_1 + ":"
							+ e.getMessage());
			throw e;
		}

		de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile tFileExcelWorkbookSave_1 = new de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile();
		// set the workbook
		tFileExcelWorkbookSave_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_2"));
		// delete template sheets if needed
		// persist workbook
		try {
			tFileExcelWorkbookSave_1
					.setOutputFile("/var/testdata/excel/excel_types_out.xlsx");
			tFileExcelWorkbookSave_1.createDirs();
			globalMap.put("tFileExcelWorkbookSave_1_COUNT_SHEETS",
					tFileExcelWorkbookSave_1.getWorkbook()
							.getNumberOfSheets());
			tFileExcelWorkbookSave_1.writeWorkbook();
			// release the memory
			globalMap.put("tFileExcelWorkbookSave_1_FILENAME",
					"/var/testdata/excel/excel_types_out.xlsx");
			globalMap.remove("workbook_tFileExcelWorkbookOpen_2");
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookSave_1_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

	}
	
	public static void testTables() throws Exception {
		de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile tFileExcelWorkbookOpen_2 = new de.cimt.talendcomp.tfileexcelpoi.SpreadsheetFile();
		tFileExcelWorkbookOpen_2.setCreateStreamingXMLWorkbook(false);
		try {
			// create empty XLSX workbook
			tFileExcelWorkbookOpen_2.setInputFile("/private/var/testdata/excel/excel_table_result.xlsx");
			tFileExcelWorkbookOpen_2.initializeWorkbook();
		} catch (Exception e) {
			globalMap.put("tFileExcelWorkbookOpen_2_ERROR_MESSAGE",
					e.getMessage());
			throw e;
		}

		globalMap.put("workbook_tFileExcelWorkbookOpen_2",
				tFileExcelWorkbookOpen_2.getWorkbook());
		globalMap.put("tFileExcelWorkbookOpen_2_COUNT_SHEETS",
				tFileExcelWorkbookOpen_2.getWorkbook()
						.getNumberOfSheets());

		int nb_line_tFileExcelSheetOutput_1 = 0; 
		de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput tFileExcelSheetOutput_1 = new de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput();
		tFileExcelSheetOutput_1
				.setWorkbook((org.apache.poi.ss.usermodel.Workbook) globalMap
						.get("workbook_tFileExcelWorkbookOpen_2"));
		tFileExcelSheetOutput_1.setTargetSheetName("Blatt1");
		tFileExcelSheetOutput_1.initializeSheet();
		Map<Integer, Integer> columnIndexes = new HashMap<Integer, Integer>();
		columnIndexes.put(0, 0);
		columnIndexes.put(1, 2);
		boolean individualColumnMappingUsed = true;
		int firstDataRowIndex = 0;
		List<XSSFTable> listTables =  ((XSSFSheet) tFileExcelSheetOutput_1.getSheet()).getTables();
		if (individualColumnMappingUsed) {
			for (Integer col : columnIndexes.values()) {
				System.out.println("check column:" + col);
				// walk through all written columns and ...
				for (int i = listTables.size() - 1; i >= 0; i--) {
					System.out.println("check table index:" + i);
					// ... through all tables
					XSSFTable table = listTables.get(i);
					// check if the table is written
					if (tFileExcelSheetOutput_1.extendTable(table, firstDataRowIndex, col.intValue(), 7)) {
						// if extended, remove it from the list
						System.out.println("table extended for column:" + col.intValue());
						listTables.remove(table);
					}
				}
			}
		} else {
			for (int i = listTables.size() - 1; i >= 0; i--) {
				// walk through all tables
				XSSFTable table = listTables.get(i);
				// check if the table is written
				if (tFileExcelSheetOutput_1.extendTable(table, firstDataRowIndex, 0, 7)) {
					// if extended, remove it from the list
					listTables.remove(table);
				}
			}
		}
		
//		XSSFSheet xs = (XSSFSheet) tFileExcelSheetOutput_1.getSheet();
//		List<XSSFTable> listTables =  xs.getTables();
//		int firstRow = 0;
//		int lastRow = 7;
//		int firstCol = 0;
//		for (XSSFTable table : listTables) {
//			AreaReference currentRef = new AreaReference(table.getCTTable().getRef());
//			CellReference topLeft = currentRef.getFirstCell();
//			CellReference buttomRight = currentRef.getLastCell();
//			if (topLeft.getRow() <= firstRow && buttomRight.getRow() >= firstRow && topLeft.getCol() >= firstCol && buttomRight.getCol() >= firstCol) {
//				// this table is within out write area, we have to expand it
//				AreaReference newRef = new AreaReference(
//						topLeft, // left top including the header line
//						new CellReference(lastRow, buttomRight.getCol())); // bottom right
//				table.getCTTable().setRef(newRef.formatAsString());
//			}
//		}
		tFileExcelWorkbookOpen_2.setOutputFile("/private/var/testdata/excel/excel_table_result_updated.xlsx");
		tFileExcelWorkbookOpen_2.evaluateAllFormulars();
		tFileExcelWorkbookOpen_2.writeWorkbook();
		
	}
	
	
}
