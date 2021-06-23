
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;

import de.jlo.talendcomp.excel.SpreadsheetInput;
import de.jlo.talendcomp.excel.SpreadsheetNamedCellInput;
import de.jlo.talendcomp.excel.SpreadsheetOutput;
import de.jlo.talendcomp.excel.SpreadsheetReferencedCellInput;


public class TestExcelComp {

	/**
	 * @param args
	 * @throws ParseException 
	 */
	public static void main(String[] args) throws Exception {
		//testReferencedCells();
//		testOutputStyled();
//		testEncryptFile();
		String s = "01:00:00 PM";
		SimpleDateFormat sdf = new SimpleDateFormat("hh:mm:ss aaa");
		System.out.println(sdf.parse(s).getTime());
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
	
	public static void test2() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		try {
			e.setInputFile("/Volumes/Data/projects/vhv/Sonderthemen_20121018.xlsx", true);
			e.initializeWorkbook();
			e.useSheet(0, false);
			e.setRowStartIndex(3);
			while (e.readNextRow()) {
				Double s = e.getDoubleCellValue(0, true, false);
				System.out.println(s);
				System.out.println("|");
				String b = e.getStringCellValue(1, true, false, false);
				System.out.println(b);
				System.out.println("|");
				Date d = e.getDateCellValue(7, true, false, "dd.MM.yyyy");
				System.out.println(d);
				System.out.println("|");
				String d1 = e.getStringCellValue(6, true, false, false);
				System.out.println(d1);
				System.out.println("-------");
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public static void test3() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		try {
			e.setInputFile("/home/jlolling/test/excel/macro_test.xlsm", true);
			e.initializeWorkbook();
			e.useSheet(1, false);
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

	public static void test1() throws ParseException {
		SpreadsheetInput e = new SpreadsheetInput();
		try {
			e.setInputFile("/var/testdata/test/excel/excel_output_file.xls", true);
			e.initializeWorkbook();
			e.useSheet(0, false);
			e.setRowStartIndex(1);
			e.setDataColumnPosition(0, "A");
			e.setDataColumnPosition(1, "B");
			e.setDataColumnPosition(2, "D");
			e.setDataColumnPosition(3, "E");
			while (e.readNextRow()) {
				String v = e.getFormularCellValue(3, true);
				Double d = e.getDoubleCellValue(3, true, false);
				System.out.println("F=" + v + " d=" + d);
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		System.out.println("lastRowNum=" + e.getLastRowNum());
	}
	
	public static void test5() throws ParseException {
		SpreadsheetInput e = new SpreadsheetInput();
		try {
			e.setInputFile("/Users/jan/test/excel/header_columns.xls", true);
			e.initializeWorkbook();
			e.useSheet(0, false);
			e.setRowStartIndex(1);
			e.setHeaderRowIndex(0);
			e.setHeaderName(0, "F0", false);
			e.setHeaderName(1, "F1", false);
			e.setHeaderName(2, "F2", true);
			e.setHeaderName(3, "F3", true);
			e.configColumnPositions();
			while (e.readNextRow()) {
				String f0 = e.getStringCellValue(0, false, true, false);
				Double f1 = e.getDoubleCellValue(1, true, false);
				String f2 = e.getStringCellValue(2, true, true, false);
				String f3 = e.getStringCellValue(3, true, true, false);
				System.out.println("F0=" + f0 + " f1=" + f1 + " f2=" + f2 + " f3=" + f3);
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		System.out.println("lastRowNum=" + e.getLastRowNum());
	}

	public static void testMassiveOutput() throws Exception {
		SpreadsheetOutput out = new SpreadsheetOutput();
		//out.setCreateStreamingXMLWorkbook(true);
		out.createEmptyXLSWorkbook();
		out.initializeWorkbook();
		out.resetCache();
		out.freezeAt(0, 1);
		for (int r = 0; r < 9; r++) {
			Object[] row = new Object[20];
			for (int c = 0; c < 2; c++) {
				row[c] = "!value:" + r + "-" + c;
			}
			out.writeRow(row);
		}
		out.setOutputFile("/var/testdata/excel/smallfile.xls");
		out.writeWorkbook();
	}
	
	public static void testEmpty() throws Exception {
		SpreadsheetOutput out = new SpreadsheetOutput();
		out.createEmptyXLSXWorkbook();
		out.initializeWorkbook();
		out.resetCache();
		System.out.println("is empty:" + out.isEmpty());
	}

	public static void testXLSDecrypted() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		e.setInputFile("/private/var/testdata/excel/encrypted.xls", true);
		e.setPassword("lolli");
		e.initializeWorkbook();
		e.useSheet(0, false);
		e.setRowStartIndex(0);
		while (e.readNextRow()) {
			System.out.println(e.getStringCellValue(0, true, true, false));
		}
	}

	public static void testXLSXDecrypted() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		e.setInputFile("/private/var/testdata/excel/encrypted.xlsx", true);
		e.setPassword("lolli");
		e.initializeWorkbook();
		e.useSheet(0, false);
		e.setRowStartIndex(0);
		while (e.readNextRow()) {
			System.out.println(e.getStringCellValue(0, true, true, false));
		}
	}

	public static void testNamedCells() throws Exception {
		SpreadsheetOutput e = new SpreadsheetOutput();
		e.setInputFile("/Volumes/Data/Talend/testdata/excel/named_cells_example.xlsx", true);
		e.initializeWorkbook();
		String name = "name_xfd";
		if (e.writeNamedCellValue(name, 1234) == false) {
			throw new Exception("cell " + name + " cannot be found");
		}
		e.setOutputFile("/Volumes/Data/Talend/testdata/excel/named_cells_example_result.xlsx");
		e.writeWorkbook();
		SpreadsheetNamedCellInput input = new SpreadsheetNamedCellInput();
		input.setInputFile("/Volumes/Data/Talend/testdata/excel/named_cells_example_result.xlsx", true);
		input.initializeWorkbook();
		input.retrieveNamedCellCount();
		System.out.println("Number of names: " + input.getNumberOfNamedCells());
		while (input.readNextNamedCell()) {
			System.out.println("------------------------------------------");
			System.out.println("Sheet:       " + input.getCellSheetName());
			System.out.println("Cell name:   " + input.getCellName());
			System.out.println("Cell address:" + input.getCellExcelReference());
			System.out.println("Cell row:    " + input.getCellRowIndex());
			System.out.println("Cell column: " + input.getCellColumnIndex());
			System.out.println("Cell value:  " + input.getCellValue());
		}
	}

	public static void testIfErrorFunction() throws Exception {
		SpreadsheetInput e = new SpreadsheetInput();
		SpreadsheetInput.registerFunction("IFERROR", "org.apache.poi.ss.formula.atp.IfError");
		e.setInputFile("/private/var/testdata/excel/irerror_test.xlsx", true);
		e.initializeWorkbook();
		e.useSheet(0, false);
		e.setRowStartIndex(0);
		while (e.readNextRow()) {
			System.out.println(e.getDoubleCellValue(0, true, false));
		}
	}

	public static void testReferencedCells() throws Exception {
		SpreadsheetReferencedCellInput e = new SpreadsheetReferencedCellInput();
		e.setInputFile("/private/var/testdata/excel/excel_sample_with_comments.xlsx", true);
		e.initializeWorkbook();
		if (e.readNextCell("Sheet1!B4", null, null, null)) {
			Cell cell = e.getCurrentCell();
			CellStyle style = cell.getCellStyle();
			System.out.println(style.getIndex());
			System.out.println(SpreadsheetReferencedCellInput.getColorString(style.getFillForegroundColorColor()));
			XSSFColor color = (XSSFColor) style.getFillBackgroundColorColor();
			System.out.println(color.getCTColor().isSetIndexed());
			System.out.println(color.getCTColor().getIndexed());
			System.out.println(HSSFColor.getIndexHash().get(color.getCTColor().getIndexed()));
			System.out.println(SpreadsheetReferencedCellInput.getColorString(color));

		} else {
			System.err.println("Cell not found");
		}
	}
	
	private static final Pattern CELL_REF_PATTERN = Pattern.compile("\\$?([A-Za-z]+)\\$?([0-9]+)");
	
	public static void testNamedCellToRange() {
		String cellFormula = /*"'sheet'!$A$1"*/ "'sheet'!$A$1:$C$4";
	    String[] refParts = cellFormula.split("!");
	    if (refParts.length == 2) {
		    String nameSheet = refParts[0].replace('\'',' ').trim();
		    if (nameSheet == null || nameSheet.isEmpty()) {
		    	return ;
		    }
	    	String cellRef = refParts[1];
		    Matcher m = CELL_REF_PATTERN.matcher(cellRef);
		    if (m.matches()) {
		    	// only allow names to single cells
		    	CellReference cellReference = new CellReference(cellRef);
			    int numRow = cellReference.getRow();
			    int numCol = cellReference.getCol();
			    System.out.println(numRow + ":" + numCol);
		    } else {
		    	System.err.println("Range cell ref not allowed");
		    }
	    } else {
	    	throw new IllegalStateException("Invalid cell reference:" + cellFormula);
	    }
	}

}
