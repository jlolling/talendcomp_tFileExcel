package de.cimt.talendcomp.tfileexcelpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFOptimiser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadsheetFile {
	
	public static enum SpreadsheetTyp {XLS, XLSX};
	protected SpreadsheetTyp currentType;
	protected FileInputStream fin;
	protected FileOutputStream fout;
	protected File outputFile;
	protected Workbook workbook;
	protected DataFormat format;
	protected Sheet sheet;
	protected String targetSheetName;
	protected Row currentRow;
	protected int currentDatasetNumber = 0;
	protected String dateFormatPattern = "dd.MM.yyyy HH:mm:ss";
	protected CreationHelper creationHelper;
	protected CellStyle cellDateStyle;
	protected int columnStartIndex = 0;
	protected boolean individualColumnMappingUsed = false;
	protected Map<Integer, Integer> columnIndexes = new HashMap<Integer, Integer>();
	protected int rowStartIndex = 0;
	protected int sheetLastRowIndex = 0;
	private static boolean functionsRegistered = false;
	private boolean createStreamingXMLWorkbook = false;
	private int rowAccessWindow = 100;
	private String password;
	private FormulaEvaluator formulaEvaluator;
	protected Map<String, CellStyle> namedStyles = new HashMap<String, CellStyle>();
	private DataFormatter dataFormatter = null;
	protected boolean debug = false;
	private static final Pattern CELL_REF_PATTERN = Pattern.compile("\\$?([A-Za-z]+)\\$?([0-9]+)"); // max is XFD

	protected DataFormatter getDataFormatter() {
		if (dataFormatter == null) {
			if (workbook instanceof HSSFWorkbook) {
				dataFormatter = new HSSFDataFormatter();
			} else {
				dataFormatter = new DataFormatter();
			}
		}
		return dataFormatter;
	}
	
	protected FormulaEvaluator getFormulaEvaluator() {
		if (formulaEvaluator == null) {
			formulaEvaluator = creationHelper.createFormulaEvaluator();
		}
		return formulaEvaluator;
	}

	public void evaluateAllFormulars() {
		getFormulaEvaluator().evaluateAll();
	}
	
	public boolean isCreateStreamingXMLWorkbook() {
		return createStreamingXMLWorkbook;
	}

	public void setCreateStreamingXMLWorkbook(boolean createStreamingXMLWorkbook) {
		this.createStreamingXMLWorkbook = createStreamingXMLWorkbook;
	}

	public SpreadsheetFile() {
		if (functionsRegistered == false) {
			functionsRegistered = true;
		}
	}
	
	public static void registerBackportFunctions() {
		try {
			//registerFunction("IFERROR", "de.cimt.talendcomp.tfileexcelpoi.functions.IfError");
		} catch (Exception e) {
			System.err.println(e.getMessage());
		}
	}
	
	/**
	 * register a function to POI
	 * @param name function name (use the English language name of the function!)
	 * @param functionClassName class name of the function including package
	 * @throws Exception if function cannot be loaded or instantiated
	 */
	public static void registerFunction(String name, String functionClassName) throws Exception {
		try {
			Object o = Class.forName(functionClassName).newInstance();
			if (o instanceof Function) {
				Function f = (Function) o;
				WorkbookEvaluator.registerFunction(name, f);
			} else if (o instanceof FreeRefFunction) {
				FreeRefFunction f = (FreeRefFunction) o;
				WorkbookEvaluator.registerFunction(name, f);
			} else {
				throw new IllegalArgumentException("Register function: " + name + " failed: Class " + functionClassName + " is not a Function or a FreeRefFunction");
			}
		} catch (ClassNotFoundException cnf) {
			throw new Exception("Register function name=" + name + " functionClassName=" + functionClassName + " failed:" + cnf.getMessage(), cnf);
		}
	}
	
	/**
	 * map the value in the row column to an excel column
	 * @param schemaColumnIndex index in the data row (parameter of writeRow method)
	 * @param columnName 'A' or 'BC'
	 */
	public void setDataColumnPosition(int schemaColumnIndex, String columnName) {
		if (columnName != null) {
			columnIndexes.put(schemaColumnIndex, CellReference.convertColStringToIndex(columnName));
			individualColumnMappingUsed = true;
		}
	}
	
	/**
	 * map the value in the row column to an excel column
	 * @param schemaColumnIndex index in the data row (parameter of writeRow method)
	 * @param sheetColumnIndex 0 - n
	 */
	public void setDataColumnPosition(int schemaColumnIndex, Integer sheetColumnIndex) {
		if (sheetColumnIndex != null) {
			columnIndexes.put(schemaColumnIndex, sheetColumnIndex);
			individualColumnMappingUsed = true;
		}
	}

	public void setTargetSheetName(String name) throws IOException {
		this.targetSheetName = ensureCorrectExcelSheetName(name);
	}
	
	public void setOutputFile(String file) throws Exception {
		try {
			SpreadsheetTyp type = getSpreadsheetType(file);
			if (type != SpreadsheetTyp.XLS && workbook instanceof HSSFWorkbook) {
				throw new Exception("Given output file name does not fit to the created workbook type xls.\nYou could leaf the extension out, the extension .xls will be added automatically.");
			} else if (type != SpreadsheetTyp.XLSX && (workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook)) {
				throw new Exception("Given output file name does not fit to the created workbook type xlsx.\nYou could leaf the extension out, the extension .xlsx will be added automatically.");
			}
		} catch (Exception e) {
			if (workbook instanceof HSSFWorkbook) {
				file = file + ".xls";
			} else {
				file = file + ".xlsx";
			}
		}
		File of = new File(file);
		setOutputFile(of);
	}
	
	public String getOutputFile() {
		return outputFile.getAbsolutePath();
	}
	
	public void setOutputFile(File outputFile) throws Exception {
		this.outputFile = outputFile;
		SpreadsheetTyp type = getSpreadsheetType(outputFile.getName());
		if (currentType != null && type != null) {
			if (currentType != type) {
				throw new Exception("Workbook cannot be saved into a type different from input type (" + currentType + ")");
			}
		}
		currentType = type;
	}

	public void createDirs() throws Exception {
		File dir = outputFile.getParentFile();
		if (dir != null) {
			dir.mkdirs();
		} else {
			throw new Exception("Output file: " + outputFile.getPath() + " has no absolute path!");
		}
	}
	
	public void createEmptyXLSWorkbook() throws IOException {
		currentType = SpreadsheetTyp.XLS;
		fin = null;
	}
	
	public void createEmptyXLSXWorkbook() throws IOException {
		currentType = SpreadsheetTyp.XLSX;
		fin = null;
	}

	public void setInputFile(String inputFileName) throws Exception {
		File inputFile = new File(inputFileName);
		SpreadsheetTyp type = getSpreadsheetType(inputFile.getName());
		if (currentType != null) {
			if (currentType != type) {
				throw new Exception("Workbook cannot be saved into a different type for output");
			}
		} else {
			currentType = type;
		}
		fin = new FileInputStream(inputFile);
	}
	
	private SpreadsheetTyp getSpreadsheetType(String name) throws Exception {
		SpreadsheetTyp type = null;
		if (name.toLowerCase().endsWith(".xls")) {
			type = SpreadsheetTyp.XLS;
		} else if (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xlsm") || name.toLowerCase().endsWith(".xlsb")) {
			type = SpreadsheetTyp.XLSX;
		} else {
			throw new Exception("Unknown type of file " + name + ". Currently supported are: xls, xlsx, xlsm, xlsb");
		}
		return type;
	}
	
	public void initializeWorkbook() throws Exception {
		if (fin != null) {
			// open existing files
			if (currentType == SpreadsheetTyp.XLS) {
				if (password != null) {
					try {
						// switch on decryption
						Biff8EncryptionKey.setCurrentUserPassword(password);
						workbook = new HSSFWorkbook(fin);
					} finally {
						// switch off
						Biff8EncryptionKey.setCurrentUserPassword(null);
						password = null;
					}
				} else {
					workbook = new HSSFWorkbook(fin);
				}
			} else if (currentType == SpreadsheetTyp.XLSX) {
				if (createStreamingXMLWorkbook) {
					workbook = new SXSSFWorkbook(new XSSFWorkbook(fin), rowAccessWindow);
				} else {
					if (password != null) {
						POIFSFileSystem filesystem = new POIFSFileSystem(fin);
						EncryptionInfo info = new EncryptionInfo(filesystem);
						Decryptor d = Decryptor.getInstance(info);
						try {
						    if (!d.verifyPassword(password)) {
						        throw new Exception("Unable to process: document is encrypted and given password does not match!");
						    }
						    // decrypt 
						    InputStream dataStream = d.getDataStream(filesystem);
						    // use open input stream
							workbook = new XSSFWorkbook(dataStream);
						} catch (GeneralSecurityException ex) {
						    throw new Exception("Unable to process encrypted document", ex);
						}
						password = null;
					} else {
						workbook = new XSSFWorkbook(fin);
					}
				}
			}
		} else {
			// create new workbooks
			if (currentType == SpreadsheetTyp.XLS) {
				workbook = new HSSFWorkbook();
			} else if (currentType == SpreadsheetTyp.XLSX) {
				if (createStreamingXMLWorkbook) {
					workbook = new SXSSFWorkbook(new XSSFWorkbook(), rowAccessWindow);
				} else {
					workbook = new XSSFWorkbook();
				}
			}
		}
		setupDataFormatStyle();
	}
	
	public int getRowAccessWindow() {
		return rowAccessWindow;
	}

	public void setRowAccessWindow(int rowAccessWindow) {
		this.rowAccessWindow = rowAccessWindow;
	}

	public void setWorkbook(Workbook wb) {
		if (wb == null) {
			throw new IllegalArgumentException("workbook cannot be null!");
		}
		this.workbook = wb;
		if (workbook instanceof HSSFWorkbook) {
			currentType = SpreadsheetTyp.XLS;
		} else if (workbook instanceof XSSFWorkbook) {
			currentType = SpreadsheetTyp.XLSX;
		} else if (workbook instanceof SXSSFWorkbook) {
			currentType = SpreadsheetTyp.XLSX;
		} else {
			throw new IllegalArgumentException("Unknown workbook type: " + workbook.getClass().getName());
		}
		setupDataFormatStyle();
	}
	
	public Workbook getWorkbook() {
		return workbook;
	}
	
	private void setupDataFormatStyle() {
		if (workbook == null) {
			throw new IllegalStateException("workbook must be initialized!");
		}
		creationHelper = workbook.getCreationHelper();
		format = workbook.createDataFormat();
		cellDateStyle = workbook.createCellStyle();
		cellDateStyle.setDataFormat(format.getFormat(dateFormatPattern));
	}
	
	public void resetDatasetNumber() {
		currentDatasetNumber = 0;
	}
	
	public int getLastRowNum() {
		if (sheet == null) {
			throw new IllegalStateException("call initializeSheet before!");
		}
		return sheet.getLastRowNum();
	}
	
	protected boolean writeInExistingCellAllowed() {
		return currentType == SpreadsheetTyp.XLSX && createStreamingXMLWorkbook;
	}
	
	protected Row getRow(int index) {
		Row row = sheet.getRow(index);
		if (row == null) {
			row = sheet.createRow(index);
		}
		return row;
	}

	protected Cell getCell(Row row, int cellIndex) {
		Cell cell = row.getCell(cellIndex);
		if (cell == null) {
			cell = row.createCell(cellIndex);
		}
		return cell;
	}
	
	protected int getCellIndex(int columnIndex) {
		if (individualColumnMappingUsed) {
			Integer cellIndex = columnIndexes.get(columnIndex);
			if (cellIndex == null) {
				cellIndex = columnIndex;
			}
			return cellIndex;
		} else {
			return columnIndex + columnStartIndex;
		}
	}
	
	protected String getFormular(String formular, int rowIndex) {
		if (formular.startsWith("=")) {
			formular = formular.substring(1);
		}
		StringReplacer sr = new StringReplacer(formular);
		sr.replace("{row}", "" + (rowIndex + 1));
		return sr.getResultText();
	}
	
	public int getLineCount() {
		return currentDatasetNumber;
	}

	public void writeWorkbook() throws Exception {
		fout = new FileOutputStream(outputFile);
        File pFile = outputFile.getParentFile();
        if (pFile != null && pFile.exists() == false) {
            pFile.mkdirs();
        }
		workbook.write(fout);
		fout.flush();
		fout.close();
		workbook = null;
	}
	
	public void deleteSheet(int sheetIndex) throws Exception {
		try {
			workbook.removeSheetAt(sheetIndex);
		} catch (Throwable t) {
			if (workbook instanceof SXSSFWorkbook) {
				throw new Exception("Deleting a sheet cannot work in a workbook which is not fully loaded because of the memory saving mode. Uncheck Memory saving mode in tFileExcelWorkbookOpen!");
			} else {
				throw new Exception("Delete sheet failed:" + t.getMessage(), t);
			}
		}
	}
	
	public void deleteSheet(String sheetName) throws Exception {
		int index = workbook.getSheetIndex(sheetName);
		if (index >= 0) {
			deleteSheet(index);
		} else {
			throw new Exception("delete sheet:" + sheetName + " failed: sheet does not exists");
		}
	}

	public int getColumnStartIndex() {
		return columnStartIndex;
	}

	public void setColumnStart(int columnStartIndex) {
		this.columnStartIndex = columnStartIndex;
	}

	public void setColumnStart(String columnName) {
		this.columnStartIndex = CellReference.convertColStringToIndex(columnName);
	}

	public int getRowStartIndex() {
		return rowStartIndex;
	}

	/**
	 * set the start row
	 * @param rowStartIndex starts with 0
	 */
	public void setRowStartIndex(int rowStartIndex) {
		this.rowStartIndex = rowStartIndex;
	}
	
    public static String ensureCorrectExcelSheetName(String desiredName) {
        if (desiredName == null || desiredName.length() == 0) {
            return "Tabelle";
        } else {
            StringReplacer sr = new StringReplacer(desiredName);
            sr.replace("/", " ");
            sr.replace("\\", " ");
            sr.replace("?", "");
            sr.replace("*", "");
            sr.replace("]", " ");
            sr.replace("[", " ");
            String newName = sr.getResultText();
            if (newName.length() > 31) {
                newName = newName.substring(0, 30);
            }
            return newName;
        }
    }

	public int getSheetLastRowIndex() {
		if (currentRow != null) {
			return currentRow.getRowNum();
		}
		return sheetLastRowIndex;
	}

	public String getDateFormatPattern() {
		return dateFormatPattern;
	}

	public void setDateFormatPattern(String dateFormatPattern) {
		if (dateFormatPattern != null && dateFormatPattern.trim().length() > 0) {
			this.dateFormatPattern = dateFormatPattern;
		}
	}
	
	public void optimizeHSSFWorkbookStyles() {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized.");
		}
		if (workbook instanceof HSSFWorkbook) {
			HSSFOptimiser.optimiseCellStyles((HSSFWorkbook) workbook);
		}
	}

	public void optimizeHSSFWorkbookFonts() {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized.");
		}
		if (workbook instanceof HSSFWorkbook) {
			HSSFOptimiser.optimiseFonts((HSSFWorkbook) workbook);
		}
	}
	
	public void setPassword(String password) {
		this.password = password;
	}
	
	protected Cell getNamedCell(String name) throws Exception {
	    Name namedCellRef = workbook.getName(name);
	    return getNamedCell(namedCellRef);
	}
	
	protected Cell getNamedCell(Name namedCellRef) {
	    if (namedCellRef != null ) {
		    String cellFormula = namedCellRef.getRefersToFormula();
		    return getCellByFormula(cellFormula);
	    } else {
	    	return null;
	    }
	}
	
	public Cell getCellByFormula(String cellFormula) {
	    if (cellFormula == null || cellFormula.isEmpty()) {
	    	return null;
	    }
	    String[] refParts = cellFormula.split("!");
	    if (refParts.length == 2) {
		    String nameSheet = refParts[0].replace('\'',' ').trim();
		    if (nameSheet == null || nameSheet.isEmpty()) {
		    	return null;
		    }
	    	String cellRef = refParts[1];
		    Matcher m = CELL_REF_PATTERN.matcher(cellRef);
		    if (m.matches()) {
		    	// only allow names refer to single cells
		    	CellReference cellReference = new CellReference(cellRef);
			    int numRow = cellReference.getRow() + 1;
			    int numCol = cellReference.getCol();
			    return getCell(nameSheet, numRow, numCol);
		    } else {
		    	return null;
		    }
	    } else {
	    	throw new IllegalStateException("Invalid cell reference:" + cellFormula);
	    }
	}
	
	private Cell getCell(String nameSheet, int numRow, int numCol) {
	    Sheet cellsheet = workbook.getSheet(nameSheet);
	    if (cellsheet == null) {
	    	return null;
	    }
	    Row row = cellsheet.getRow(numRow - 1);
	    if (row == null) {
	    	row = cellsheet.createRow(numRow - 1);
	    }
	    Cell cell = row.getCell(numCol);
	    if (cell == null) {
	    	cell = row.createCell(numCol);
	    }
	    return cell;
	}
	
	protected Cell getReferencedCell(String cellRefStr) throws Exception {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized.");
		}
		CellReference cellRef = new CellReference(cellRefStr);
		String sheetName = cellRef.getSheetName();
		Sheet cellsheet = null;
		if (sheetName != null && sheetName.isEmpty() == false) {
			if (sheet == null || sheet.getSheetName().equalsIgnoreCase(sheetName) == false) {
				cellsheet = workbook.getSheet(sheetName);
			} else {
				throw new Exception("Sheet with name:" + sheetName + " does not exists.");
			}
		} else {
			if (sheet == null) {
				throw new Exception("No current sheet selected. The given cell reference:" + cellRefStr + " contains not sheet name.");
			} else {
				cellsheet = sheet;
			}
		}
		if (cellsheet != null) {
			int numRow = cellRef.getRow();
			Row row = cellsheet.getRow(numRow);
			if (row == null) {
				row = cellsheet.createRow(numRow);
			}
			short numCol = cellRef.getCol();
			Cell cell = row.getCell(numCol);
			if (cell == null) {
				cell = row.createCell(numCol);
			}
			return cell;
		} else {
			return null;
		}
	}
	
	/**
	 * adds a font to the workbook
	 * @param family like Arial
	 * @param height like 8,9,10,12,14...
	 * @param bui with "b"=bold, "u"=underlined, "i"=italic and all combinations as String
	 * @param color color index
	 */
	public void addStyle(String styleName, String fontFamily, String fontHeight, String fontDecoration, String fontColor, String bgColor, String textAlign, boolean buttomBorder) {
		if (styleName != null && styleName.isEmpty() == false) {
			Font f = workbook.createFont();
			if (fontFamily != null && fontFamily.isEmpty() == false) {
				f.setFontName(fontFamily);
			}
			if (fontHeight != null && fontHeight.isEmpty() == false) {
				short height = Short.parseShort(fontHeight);
				if (height > 0) {
					f.setFontHeightInPoints(height);
				}
			}
			if (fontDecoration != null && fontDecoration.isEmpty() == false) {
				if (fontDecoration.contains("b")) {
					f.setBoldweight(Font.BOLDWEIGHT_BOLD);
				}
				if (fontDecoration.contains("i")) {
					f.setItalic(true);
				}
				if (fontDecoration.contains("u")) {
					f.setUnderline(Font.U_SINGLE);
				}
			}
			if (fontColor != null && fontColor.isEmpty() == false) {
				short color = Short.parseShort(fontColor);
				f.setColor(color);
			}
			CellStyle style = workbook.createCellStyle();
			style.setFont(f);
			if (bgColor != null && bgColor.isEmpty() == false) {
				short color = Short.parseShort(bgColor);
				style.setFillForegroundColor(color);
				//style.setFillBackgroundColor(color);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			}
			if (textAlign != null && textAlign.isEmpty() == false) {
				if ("center".equalsIgnoreCase(textAlign)) {
					style.setAlignment(CellStyle.ALIGN_CENTER);
				} else if ("left".equalsIgnoreCase(textAlign)) {
					style.setAlignment(CellStyle.ALIGN_LEFT);
				} else if ("right".equals(textAlign)) {
					style.setAlignment(CellStyle.ALIGN_RIGHT);
				}
			}
			if (buttomBorder) {
				style.setBorderBottom(CellStyle.BORDER_MEDIUM);
				style.setBottomBorderColor((short) 9);
			}
			namedStyles.put(styleName, style);
		}
	}
		
	public boolean isEmpty() {
		if (workbook == null) {
			throw new IllegalStateException("workbook is not initialized");
		}
		int countSheets = workbook.getNumberOfSheets();
		if (countSheets == 0) {
			return true;
		}
		for (int i = 0; i < countSheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			if (sheet.getLastRowNum() == 0) {
				return true;
			}
		}
		return false;
	}

	public boolean isDebug() {
		return debug;
	}

	public void setDebug(boolean debug) {
		this.debug = debug;
	}

}