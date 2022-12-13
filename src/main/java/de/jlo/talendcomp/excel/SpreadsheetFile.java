/**
 * Copyright 2015 Jan Lolling jan.lolling@gmail.com
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.jlo.talendcomp.excel;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFOptimiser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
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
	private File inputFile = null;
	private byte[] inputBytes = null;
	private Boolean isFileMode = null;
	protected FileOutputStream fout;
	protected File outputFile = null;
	protected Workbook workbook;
	protected DataFormat format;
	protected Sheet sheet;
	protected String targetSheetName;
	protected Row currentRow;
	protected int currentRecordIndex = 0;
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
	private String readPassword;
	private FormulaEvaluator formulaEvaluator;
	private DataFormatter dataFormatter = null;
	protected boolean debug = false;
	private static final Pattern CELL_REF_PATTERN = Pattern.compile("\\$?([A-Za-z]+)\\$?([0-9]+)"); // max is XFD
	private static byte[] xlsMagicNumbers = {
		(byte)0xD0, (byte)0xCF,
		(byte)0x11, (byte)0xE0,
		(byte)0xA1, (byte)0xB1, 
		(byte)0x1A, (byte)0xE1
	};

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
		if (workbook instanceof SXSSFWorkbook) {
			warn("Skip formula evaluation because of using of a streaming-workbook. This kind of workbooks does not provide access to all rows.");
		} else {
			getFormulaEvaluator().evaluateAll();
		}
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
			//registerFunction("IFERROR", "de.jlo.talendcomp.tfileexcelpoi.functions.IfError");
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
			Object o = Class.forName(functionClassName).getDeclaredConstructor().newInstance();
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

	public void setTargetSheetName(String name) throws Exception {
		setTargetSheetName(name, false);
	}
	
	public void setTargetSheetName(String name, boolean tolerant) throws Exception {
		name = ensureCorrectExcelSheetName(name);
		sheet = findSheet(name, tolerant);
		if (sheet == null) {
			// sheet not found, so create it just now
			this.targetSheetName = name;
			sheet = workbook.createSheet(name);
		} else {
			this.targetSheetName = sheet.getSheetName();
		}
	}
	
	public void setTargetSheetName(Integer index) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		}
		if (index == null) {
			throw new Exception("Index cannot be null!");
		}
		if (index >= workbook.getNumberOfSheets()) {
			targetSheetName = "Sheet " + index;
			sheet = workbook.createSheet(targetSheetName);
		} else {
			sheet = workbook.getSheetAt(index);
			targetSheetName = sheet.getSheetName();
		}
		if (sheet == null) {
			throw new Exception("Sheet with index:" + index + " does not exist!");
		}
	}
	
	public String getTargetSheetName() {
		return targetSheetName;
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
			ensureDirExists(outputFile);
		} else {
			throw new Exception("Output file: " + outputFile.getPath() + " has not an absolute path!");
		}
	}
	
	public void createEmptyXLSWorkbook() {
		currentType = SpreadsheetTyp.XLS;
	}
	
	public void createEmptyXLSXWorkbook() {
		currentType = SpreadsheetTyp.XLSX;
	}

	/**
	 * Set the excel input file and throw an exception if the file does not exists.
	 * @param inputFileName
	 * @throws Exception
	 */
	public void setInputFile(String inputFileName) throws Exception {
		setInputFile(inputFileName, true);
	}
	
	/**
	 * Set the excel input file
	 * @param inputFileName 
	 * @param dieIfFileNotExists
	 * @throws Exception
	 */
	public void setInputFile(String inputFileName, boolean dieIfFileNotExists) throws Exception {
		if (inputFileName == null || inputFileName.trim().isEmpty() || inputFileName.trim().length() < 5) {
			throw new Exception("Input file name cannot be null or empty and must have at least 5 letters");
		}
		SpreadsheetTyp type = getSpreadsheetType(inputFileName);
		if (currentType != null) {
			if (currentType != type) {
				throw new Exception("Workbook cannot be saved into a different type for output");
			}
		} else {
			currentType = type;
		}
		File inputFile = new File(inputFileName);
		if (inputFile.exists() == false || inputFile.canRead() == false) {
			if (dieIfFileNotExists) {
				throw new Exception("Excel file: " + inputFileName + " does not exists or canot be read!");
			} else {
				this.inputFile = null;
			}
		} else {
			this.inputFile = inputFile;
			this.isFileMode = true;
		}
	}
	
	/**
	 * Set the excel byte array
	 * @param bytes
	 * @param filetype
	 * @param dieIfEmpty
	 * @throws Exception
	 */
	public void setInputFile(byte[] bytes, boolean dieIfEmpty) throws Exception {
		
		if((bytes == null || bytes.length == 0) && dieIfEmpty) {
			throw new Exception("No bytes where given as input!");
		}
		
		SpreadsheetTyp type = null;
		
		if( Arrays.equals(xlsMagicNumbers, Arrays.copyOf(bytes, xlsMagicNumbers.length)) ) {
			type = SpreadsheetTyp.XLS;
		} else {
			type = SpreadsheetTyp.XLSX;
		}
		
		if (currentType != null) {
			if (currentType != type) {
				throw new Exception("Workbook cannot be saved into a different type for output");
			}
		} else {
			currentType = type;
		}
		
		this.inputBytes = bytes;
		this.isFileMode = false;
	}
	
	private static SpreadsheetTyp getSpreadsheetType(String name) throws Exception {
		SpreadsheetTyp type = null;
		if (name.toLowerCase().endsWith(".xls")) {
			type = SpreadsheetTyp.XLS;
		} else if (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xlsm") || name.toLowerCase().endsWith(".xlsb")) {
			type = SpreadsheetTyp.XLSX;
		} else {
			throw new Exception("Unknown or missing type of the file " + name + ". Currently are supported: xls, xlsx, xlsm, xlsb");
		}
		return type;
	}
	
	private InputStream getInputStream() throws Exception {
		if(this.isFileMode) {
			return new FileInputStream(this.inputFile);
		} else {
			return new ByteArrayInputStream(this.inputBytes);
		}
	}
	
	@SuppressWarnings("resource")
	public void initializeWorkbook() throws Exception {
		if (inputFile != null || inputBytes != null) {
			// open existing files
			InputStream ins = null;
			
			if (currentType == SpreadsheetTyp.XLS) {
				if (readPassword != null) {
					try {
						// switch on decryption
						Biff8EncryptionKey.setCurrentUserPassword(readPassword);
						ins = getInputStream();
						workbook = new HSSFWorkbook(ins);
						ins.close();
					} finally {
						// switch off
						Biff8EncryptionKey.setCurrentUserPassword(null);
						readPassword = null;
					}
				} else {
					ins = getInputStream();
					workbook = new HSSFWorkbook(ins);
					ins.close();
				}
			} else if (currentType == SpreadsheetTyp.XLSX) {
				if (createStreamingXMLWorkbook) {
					ins = getInputStream();
					try {
						ZipSecureFile.setMinInflateRatio(0);
						workbook = new SXSSFWorkbook(new XSSFWorkbook(ins), rowAccessWindow);
					} finally {
						if (ins != null) {
							try {
								ins.close();
							} catch (IOException ioe) {
								// ignore
							}
						}
					}
				} else {
					if (readPassword != null) {
						ins = getInputStream();
						POIFSFileSystem filesystem = new POIFSFileSystem(ins);
						EncryptionInfo info = new EncryptionInfo(filesystem);
						Decryptor d = Decryptor.getInstance(info);
						InputStream dataStream = null;
						try {
						    if (!d.verifyPassword(readPassword)) {
						        throw new Exception("Unable to process: document is encrypted and given password does not match!");
						    }
						    // decrypt 
						    dataStream = d.getDataStream(filesystem);
						    // use open input stream
							workbook = new XSSFWorkbook(dataStream);
							dataStream.close();
						} catch (GeneralSecurityException ex) {
						    throw new Exception("Unable to read and parse encrypted document", ex);
						} finally {
							if (dataStream != null) {
								try {
									dataStream.close();
								} catch (IOException ioe) {
									// ignore
								}
							}
							if (ins != null) {
								try {
									ins.close();
								} catch (IOException ioe) {
									// ignore
								}
							}
						}
						readPassword = null;
					} else {
						ins = getInputStream();
						try {
							workbook = new XSSFWorkbook(ins);
						} finally {
							if (ins != null) {
								try {
									ins.close();
								} catch (IOException ioe) {
									// ignore
								}
							}
						}
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
			} else {
				throw new IllegalStateException("Create new workbook failed: Unknown workbook type: " + currentType + ". No workbook created!");
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
		currentRecordIndex = 0;
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
		return currentRecordIndex;
	}

	public void writeWorkbook() throws Exception {
        File pFile = outputFile.getParentFile();
        if (pFile != null && pFile.exists() == false) {
            pFile.mkdirs();
            if (pFile.exists() == false) {
            	throw new Exception("Unable to create directory: " + pFile.getAbsolutePath());
            }
        }
        Exception ex = null;
        try {
			fout = new FileOutputStream(outputFile);
			workbook.write(fout);
			fout.flush();
        } catch (Exception e) {
        	ex = e;
        	error("write workbook failed: " + e.getMessage(), e);
        } finally {
        	if (fout != null) {
        		try {
        			fout.close();
        			if (workbook instanceof SXSSFWorkbook) {
        				((SXSSFWorkbook) workbook).dispose();
        			} else {
        				workbook.close();
        			}
        		} catch (Exception e1) {
        			// ignored
        		}
        	}
        }
        if (ex != null) {
        	throw ex;
        }
		workbook = null;
	}
	
	public void writeWorkbookEncrypted(String password) throws Exception {
		if (password == null || password.trim().isEmpty()) {
			throw new Exception("Unable to encrypt while writing excel file: " + outputFile.getName() + ": Password cannot be null or empty");
		}
        File pFile = outputFile.getParentFile();
        if (pFile != null && pFile.exists() == false) {
            pFile.mkdirs();
            if (pFile.exists() == false) {
            	throw new Exception("Unable to create directory: " + pFile.getAbsolutePath());
            }
        }
        Exception ex = null;
        try {
            POIFSFileSystem fs = new POIFSFileSystem();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            Encryptor enc = info.getEncryptor();
            enc.confirmPassword(password);
			// write the workbook into the encrypted OutputStream
			OutputStream encos = enc.getDataStream(fs);
			workbook.write(encos);
			workbook.close();
			encos.close(); // this is necessary before writing out the FileSystem
			fout = new FileOutputStream(outputFile);
			fs.writeFilesystem(fout);
			fout.close();
			fs.close();
        } catch (Exception e) {
        	ex = e;
        	error("write workbook failed: " + e.getMessage(), e);
        } finally {
        	if (fout != null) {
        		try {
        			fout.close();
        			if (workbook instanceof SXSSFWorkbook) {
        				((SXSSFWorkbook) workbook).dispose();
        			} else {
        				workbook.close();
        			}
        		} catch (Exception e1) {
        			// ignored
        		}
        	}
        }
        if (ex != null) {
        	throw ex;
        }
		workbook = null;
	}

	public static void ensureDirExists(File file) throws Exception {
		File dir = file.getParentFile();
		if (dir.exists() == false) {
			dir.mkdirs();
			if (dir.exists() == false) {
				throw new Exception("Directory: " + dir.getAbsolutePath() + " does not exists and cannot be created!");
			}
		}
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
		if (rowStartIndex < 0) {
			throw new IllegalArgumentException("Row index starts with 1"); // message for the Talend users!
		}
		this.rowStartIndex = rowStartIndex;
		currentRecordIndex = 0;
	}
	
    public static String ensureCorrectExcelSheetName(String desiredName) {
        if (desiredName == null || desiredName.length() == 0) {
            return "Sheet 1";
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
    
    public int getCurrentSheetLastPresentRowIndex() {
    	if (sheet != null) {
    		return sheet.getLastRowNum();
    	} else {
    		return -1;
    	}
    }

	public int detectCurrentSheetLastNoneEmptyRowIndex() {
		int lastRowNum = sheet.getLastRowNum() + 1;
		while (lastRowNum > 0) {
			Row row = sheet.getRow(lastRowNum);
			if (row == null || isRowEmpty(row)) {
				// we found a empty row, step to the previous row
				lastRowNum = lastRowNum - 1;
			} else {
				break; // we found a none empty row, thats the last!
			}
		}
		return lastRowNum;
	}
	
	protected boolean isRowEmpty(Row row) {
		for (Cell c : row) {
			if (c.getCellType() != CellType.BLANK) {
				return false;
			}
		}
		return true;
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
		this.readPassword = password;
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
	
	public void info(String message) {
		System.out.println("INFO: " + message);
	}
	
	public void debug(String message) {
		System.out.println("DEBUG: " + message);
	}

	public void warn(String message) {
		System.err.println("WARN: " + message);
	}

	public void error(String message, Exception e) {
		System.err.println("ERROR: " + message);
		if (e != null) {
			e.printStackTrace();
		}
	}

	protected String printArray(Object[] array) {
		if (array != null) {
			StringBuilder sb = new StringBuilder();
			sb.append("[");
			for (int i = 0; i < array.length; i++) {
				if (i > 0) {
					sb.append(",");
				}
				if (array[i] != null) {
					sb.append(array[i]);
				} else {
					sb.append("null");
				}
			}
			sb.append("]");
			return sb.toString();
		} else {
			return "";
		}
	}
	
	public void close() {
		if (workbook != null) {
			try {
				workbook.close();
			} catch (Exception e) {
				// ignore
			}
		}
	}
	
	public void setZipBombWarningThreshold(Number ratio) {
		if (ratio != null && ratio.doubleValue() <= 0.01d) {
			ZipSecureFile.setMinInflateRatio(ratio.doubleValue());
		}
	}

	public Row getCurrentRow() {
		return currentRow;
	}
	
	public List<Sheet> getSheets() {
		List<Sheet> sheets = new ArrayList<>();
		int n = workbook.getNumberOfSheets();
		for (int i = 0; i < n; i++) {
			sheets.add(workbook.getSheetAt(i));
		}
		return sheets;
	}
	
	private String crunchSheetName(String name) {
		return name.toLowerCase().replace("_", "").replace(" ", "").replace("-", "").trim();
	}
	
	public Sheet findSheet(String expectedSheetName, boolean tolerant) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		}
		if (expectedSheetName == null || expectedSheetName.trim().isEmpty()) {
			throw new Exception("Name of sheet cannot be null or empty!");
		}
		Sheet expectedSheet = workbook.getSheet(expectedSheetName);
		if (expectedSheet == null && tolerant) {
			List<Sheet> sheets = getSheets();
			String crunchedExpectedSheetName = crunchSheetName(expectedSheetName);
			for (Sheet sheetInList : sheets) {
				String name = sheetInList.getSheetName();
				if (name != null) {
					String crunchedSheetNameInList = crunchSheetName(name);
					if (crunchedSheetNameInList.equals(crunchedExpectedSheetName)) {
						targetSheetName = expectedSheetName;
						expectedSheet = sheetInList;
						System.out.println("Found with tolerance the actual existing sheet: " + sheetInList.getSheetName() + ". Given expected sheet name was: " + expectedSheetName);
					}
				}
			}
		}
		return expectedSheet;
	}
	
}
