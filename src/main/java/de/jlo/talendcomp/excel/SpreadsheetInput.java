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

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;

import de.jlo.talendcomp.excel.GenericDateUtil.DateParser;

public class SpreadsheetInput extends SpreadsheetFile {
	
	private Map<Integer, Object> lastValueMap = new HashMap<Integer, Object>();
	private SimpleDateFormat defaultDateFormat = null;
	private int maxRowIndex = 0;
	private int currentRowIndex = 0;
	private NumberFormat defaultNumberFormat = null;
	private Map<Integer, NumberFormat> numberFormatColumnMap = new HashMap<Integer, NumberFormat>();
	private int headerRowIndex = 0;
	private Map<String, Integer> namesFromHeaderRow = new HashMap<String, Integer>();
	private Map<Integer, String> namesFromSchema = new HashMap<Integer, String>();
	private Map<Integer, Boolean> ignoreMissingMap = new HashMap<Integer, Boolean>();
	private Set<Integer> missingColumns = new TreeSet<Integer>();
	private Row headerRow;
	private boolean returnURLInsteadOfName = false;
	private boolean concatenateLabelUrl = false;
	private boolean findHeaderPosByRegex = false;
	private boolean useCachedValuesForFailedEvaluations = true;
	private boolean stopAtMissingRow = true;
	private StyleUtil styleUtil = null;
	private boolean overrideExcelNumberFormat = false;
	private Locale defaultLocale = null;
	private boolean parseDateFromVisibleString = false;
	private boolean lenientDateParsing = true;
	private boolean returnZeroDateAsNull = true;
	
	public SpreadsheetInput() {
		defaultNumberFormat = NumberFormat.getInstance(Locale.ENGLISH);
		defaultNumberFormat.setMaximumFractionDigits(20);
		defaultNumberFormat.setGroupingUsed(false);
	}
	
	public void setNumberPrecision(int columnIndex, Integer precision) {
		if (precision != null) {
			NumberFormat nf = (NumberFormat) defaultNumberFormat.clone();
			nf.setMaximumFractionDigits(precision);
			overrideExcelNumberFormat = true;
			numberFormatColumnMap.put(columnIndex, nf);
		}
	}
	
	private NumberFormat getNumberFormat(int columnIndex) {
		NumberFormat nf = numberFormatColumnMap.get(columnIndex);
		if (nf != null) {
			return nf;
		} else {
			return defaultNumberFormat;
		}
	}
	
	private StyleUtil getStyleUtil() {
		if (styleUtil == null) {
			styleUtil = new StyleUtil(workbook);
		}
		return styleUtil;
	}
	
	public String getCellStyleCSS(int columnIndex) {
		Cell cell = getCell(columnIndex);
		if (cell != null) {
			CellStyle style = cell.getCellStyle();
			if (style != null) {
				return getStyleUtil().buildCSS(style);
			} else {
				return "";
			}
		} else {
			CellStyle style = workbook.getCellStyleAt(0);
			if (style != null) {
				return getStyleUtil().buildCSS(style);
			} else {
				return "";
			}
		}
	}
	
	public CellStyle getCellStyle(int columnIndex) {
		Cell cell = getCell(columnIndex);
		if (cell != null) {
			return cell.getCellStyle();
		} else {
			return null;
		}
	}
	
	public String getStringCellValue(int columnIndex, boolean nullable, boolean trim, boolean useLast) throws Exception {
		String value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			value = getStringCellValue(cell, columnIndex);
		}
		if (trim && value != null) {
			value = value.trim();
		}
		if (useLast && (value == null || value.isEmpty())) {
			value = (String) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		if ((value == null || value.isEmpty()) && nullable == false) {
			throw new Exception("Cell in column " + columnIndex + " has no value!");
		}
		return value;
	}
	
	private String getStringCellValue(Cell cell, int originalColumnIndex) throws Exception {
		String value = null;
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.FORMULA) {
				try {
					value = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
				} catch (Exception e) {
					if (useCachedValuesForFailedEvaluations) {
						cellType = cell.getCachedFormulaResultTypeEnum();
						if (cellType == CellType.STRING) {
							if (returnURLInsteadOfName) {
								Hyperlink link = cell.getHyperlink();
								if (link != null) {
									if (concatenateLabelUrl) {
										String url = link.getAddress();
										if (url == null) {
											url = "";
										}
										String label = link.getLabel();
										if (label == null) {
											label = "";
										}
										value = label + "|" + url;
									} else {
										value = link.getAddress();
									}
								} else {
									value = cell.getStringCellValue();
								}
							} else {
								value = cell.getStringCellValue();
							}
						} else if (cellType == CellType.NUMERIC) {
							if (DateUtil.isCellDateFormatted(cell)) {
								if (defaultDateFormat != null) {
									Date d = cell.getDateCellValue();
									if (d != null) {
										value = defaultDateFormat.format(d);
									}
								} else {
									value = getDataFormatter().formatCellValue(cell);
								}
							} else {
								if (overrideExcelNumberFormat) {
									value = getNumberFormat(originalColumnIndex).format(cell.getNumericCellValue());
								} else {
									value = getDataFormatter().formatCellValue(cell);
								}
							}
						} else if (cellType == CellType.BOOLEAN) {
							value = cell.getBooleanCellValue() ? "true" : "false";
						}
					} else {
						throw e;
					}
				}
			} else if (cellType == CellType.STRING) {
				if (returnURLInsteadOfName) {
					Hyperlink link = cell.getHyperlink();
					if (link != null) {
						if (concatenateLabelUrl) {
							String url = link.getAddress();
							if (url == null) {
								url = "";
							}
							String label = link.getLabel();
							if (label == null) {
								label = "";
							}
							value = label + "|" + url;
						} else {
							value = link.getAddress();
						}
					} else {
						value = cell.getStringCellValue();
					}
				} else {
					value = cell.getStringCellValue();
				}
			} else if (cellType == CellType.NUMERIC) {
				if (DateUtil.isCellDateFormatted(cell)) {
					value = getDataFormatter().formatCellValue(cell);
				} else {
					if (overrideExcelNumberFormat) {
						value = getNumberFormat(originalColumnIndex).format(cell.getNumericCellValue());
					} else {
						value = getDataFormatter().formatCellValue(cell);
					}
				}
			} else if (cellType == CellType.BOOLEAN) {
				value = cell.getBooleanCellValue() ? "true" : "false";
			} else if (cellType == CellType.BLANK) {
				value = null;
			}
		}
		return value;
	}
	
	public String getFormularCellValue(int columnIndex, boolean nullable) throws Exception {
		String value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
			return null;
		} else {
			value = cell.getCellFormula();
		}
		return value;
	}

	public String getCommentCellValue(int columnIndex, boolean nullable, boolean trim, boolean useLast) throws Exception {
		String value =  null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			Comment comment = cell.getCellComment();
			if (comment == null) {
				if (nullable == false) {
					throw new Exception("Cell in column " + columnIndex + " has no value!");
				}
			} else {
				RichTextString rt = comment.getString();
				if (rt == null) {
					if (nullable == false) {
						throw new Exception("Cell in column " + columnIndex + " has no value!");
					}
				} else {
					value = rt.getString();
					if (value != null) {
						value = value.trim();
					}
				}
			}
		}
		if (useLast && (value == null || value.isEmpty())) {
			value = (String) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		return value;
	}
	
	private Cell getCell(int columnIndex) {
		if (missingColumns.contains(columnIndex)) {
			return null;
		} else {
			if (currentRow != null) {
				return currentRow.getCell(getCellIndex(columnIndex));
			} else {
				return null;
			}
		}
	}

	public boolean isCellValueEmpty(int columnIndex) {
		Cell cell = getCell(columnIndex);
		return isCellValueEmpty(cell);
	}

	public boolean isCellValueEmpty(Cell cell) {
		if (cell == null) {
			return true;
		} else { 
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.BLANK) {
				return true;
			} else if (cellType == CellType.FORMULA) {
				try {
					String s = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
					if (s == null || s.trim().isEmpty()) {
						return true;
					} else {
						return false;
					}
				} catch (Exception e) {
					return true;
				}
			} else if (cellType == CellType.STRING) {
				String s = cell.getStringCellValue();
				if (s == null || s.trim().isEmpty()) {
					return true;
				} else {
					return false;
				}
			} else {
				return false;
			}
		}
	}

	public boolean isCellCommentEmpty(int columnIndex) {
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			return true;
		}
		Comment comment = cell.getCellComment();
		if (comment == null) {
			return true;
		} else {
			RichTextString rt = comment.getString();
			if (rt == null) {
				return true;
			} else {
				return rt.getString() != null ? rt.getString().trim().isEmpty() : true;
			}
		}
	}
	
	public Double getDoubleCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			value = getDoubleCellValue(cell);
		}
		if (useLast && value == null) {
			value = (Double) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		if (value == null && nullable == false) {
			throw new Exception("Cell in column " + columnIndex + " has no value!");
		}
		return value;
	}
	
	private Double getDoubleCellValue(Cell cell) throws Exception {
		Double value = null;
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.FORMULA) {
				try {
					String s = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
					if (s != null && s.trim().isEmpty() == false) {
						Number n = getNumberFormat(cell.getColumnIndex()).parse(s.trim());
						value = n.doubleValue();
					}
				} catch (Exception e) {
					if (useCachedValuesForFailedEvaluations) {
						cellType = cell.getCachedFormulaResultTypeEnum();
						if (cellType == CellType.STRING) {
							String s = cell.getStringCellValue();
							if (s != null && s.trim().isEmpty() == false) {
								Number n = getNumberFormat(cell.getColumnIndex()).parse(s.trim());
								value = n.doubleValue();
							}
						} else if (cellType == CellType.NUMERIC) {
							value = cell.getNumericCellValue();
						}
					} else {
						throw e;
					}
				}
			} else if (cellType == CellType.STRING) {
				String s = cell.getStringCellValue();
				if (s != null && s.trim().isEmpty() == false) {
					Number n = getNumberFormat(cell.getColumnIndex()).parse(s.trim());
					value = n.doubleValue();
				}
			} else if (cellType == CellType.NUMERIC) {
				value = cell.getNumericCellValue();
			}
		}
		return value;
	}

	public BigDecimal getBigDecimalCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double d = getDoubleCellValue(columnIndex, nullable, useLast);
		if (d != null) {
			return new BigDecimal(d);
		} else {
			return null;
		}
	}

	public Integer getIntegerCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double d = getDoubleCellValue(columnIndex, nullable, useLast);
		if (d != null) {
			return new Integer(d.intValue());
		} else {
			return null;
		}
	}

	public Long getLongCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double d = getDoubleCellValue(columnIndex, nullable, useLast);
		if (d != null) {
			return new Long(d.longValue());
		} else {
			return null;
		}
	}

	public Short getShortCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double d = getDoubleCellValue(columnIndex, nullable, useLast);
		if (d != null) {
			return new Short(d.shortValue());
		} else {
			return null;
		}
	}

	public Float getFloatCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Double d = getDoubleCellValue(columnIndex, nullable, useLast);
		if (d != null) {
			return new Float(d.floatValue());
		} else {
			return null;
		}
	}

	private static Boolean toBool(String s) {
		if (s == null) {
			return null;
		}
		s = s.trim().toLowerCase();
		if ("true".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("false".equals(s)) {
			return Boolean.FALSE;
		} else if ("1".equals(s)) {
			return Boolean.TRUE;
		} else if ("0".equals(s)) {
			return Boolean.FALSE;
		} else if ("yes".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("y".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("no".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if ("n".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if ("ja".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("j".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("nein".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if ("oui".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("non".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if ("ok".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("x".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("wahr".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("falsch".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if ("vrai".equalsIgnoreCase(s)) {
			return Boolean.TRUE;
		} else if ("fausse".equalsIgnoreCase(s)) {
			return Boolean.FALSE;
		} else if (s != null) {
			return Boolean.FALSE;
		} else {
			return null;
		}
	}
	
	private Boolean toBool(Number s) {
		if (s == null) {
			return null;
		} else if (s.intValue() == 0) {
			return Boolean.FALSE;
		} else if (s.intValue() > 0) {
			return Boolean.TRUE;
		} else {
			return null;
		}
	}

	public Boolean getBooleanCellValue(int columnIndex, boolean nullable, boolean useLast) throws Exception {
		Boolean value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			value = getBooleanCellValue(cell);
		}
		if (useLast && value == null) {
			value = (Boolean) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		if (value == null && nullable == false) {
			throw new Exception("Cell in column " + columnIndex + " has no value!");
		}
		return value;
	}
	
	private Boolean getBooleanCellValue(Cell cell) throws Exception {
		Boolean value = null;
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.FORMULA) {
				try {
					String s = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
					value = toBool(s);
				} catch (Exception e) {
					if (useCachedValuesForFailedEvaluations) {
						cellType = cell.getCachedFormulaResultTypeEnum();
						if (cellType == CellType.STRING) {
							String s = cell.getStringCellValue();
							value = toBool(s);
						} else if (cellType == CellType.NUMERIC) {
							double s = cell.getNumericCellValue();
							value = toBool(s);
						} else if (cellType == CellType.BOOLEAN) {
							value = cell.getBooleanCellValue();
						}
					}
				}
			} else if (cellType == CellType.STRING) {
				String s = cell.getStringCellValue();
				value = toBool(s);
			} else if (cellType == CellType.NUMERIC) {
				double s = cell.getNumericCellValue();
				value = toBool(s);
			} else if (cellType == CellType.BOOLEAN) {
				value = cell.getBooleanCellValue();
			}
		}
		return value;
	}

	public Date getDateCellValue(int columnIndex, boolean nullable, boolean useLast, String pattern) throws Exception {
		Date value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			value = getDateCellValue(cell, pattern);
		}
		if (useLast && value == null) {
			value = (Date) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		if (value == null && nullable == false) {
			throw new Exception("Cell in column " + columnIndex + " has no value or a zero date!");
		}
		return value;
	}
	
	public Date getDurationCellValue(int columnIndex, boolean nullable, boolean useLast, String pattern) throws Exception {
		Date value = null;
		Cell cell = getCell(columnIndex);
		if (cell == null) {
			if (nullable == false) {
				throw new Exception("Cell in column " + columnIndex + " has no value!");
			}
		} else {
			value = getDurationCellValue(cell, pattern);
		}
		if (useLast && value == null) {
			value = (Date) lastValueMap.get(columnIndex);
		} else {
			lastValueMap.put(columnIndex, value);
		}
		if (value == null && nullable == false) {
			throw new Exception("Cell in column " + columnIndex + " has no value!");
		}
		return value;
	}

	private Date parseDate(String s, String pattern) throws ParseException {
		if (s != null && s.isEmpty() == false) {
			DateParser du = GenericDateUtil.getDateParser(lenientDateParsing);
			return du.parseDate(s, defaultLocale, pattern);
		} else {
			return null;
		}
	}
	
	private Date parseDuration(String s, String pattern) throws ParseException {
		if (s != null && s.isEmpty() == false) {
			return GenericDateUtil.parseDuration(s, pattern);
		} else {
			return null;
		}
	}

	private Date getDateCellValue(Cell cell, String pattern) throws Exception {
		Date value = null;
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.FORMULA) {
				try {
					String s = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
					return parseDate(s, pattern);
				} catch (Exception e) {
					if (useCachedValuesForFailedEvaluations) {
						cellType = cell.getCachedFormulaResultTypeEnum();
						if (cellType == CellType.STRING) {
							String s = cell.getStringCellValue();
							value = parseDate(s, pattern);
						} else if (cellType == CellType.NUMERIC) {
							value = cell.getDateCellValue();
						}
					} else {
						throw e;
					}
				}
			} else if (cellType == CellType.NUMERIC) {
				if (DateUtil.isCellDateFormatted(cell) && parseDateFromVisibleString == false) {
					value = cell.getDateCellValue();
				} else {
					String s = getDataFormatter().formatCellValue(cell);
					value = parseDate(s, pattern);
				}
			} else if (cellType == CellType.STRING) {
				String s = getDataFormatter().formatCellValue(cell);
				value = parseDate(s, pattern);
			}
		}
		if (returnZeroDateAsNull && GenericDateUtil.isZeroDate(value)) {
			value = null;
		}
		return value;
	}
	
	private Date getDurationCellValue(Cell cell, String pattern) throws Exception {
		Date value = null;
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType == CellType.FORMULA) {
				try {
					String s = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
					return parseDuration(s, pattern);
				} catch (Exception e) {
					if (useCachedValuesForFailedEvaluations) {
						cellType = cell.getCachedFormulaResultTypeEnum();
						if (cellType == CellType.STRING) {
							String s = getDataFormatter().formatCellValue(cell);
							value = parseDate(s, pattern);
						} else if (cellType == CellType.NUMERIC) {
							value = cell.getDateCellValue();
						}
					} else {
						throw e;
					}
				}
			} else if (cellType == CellType.NUMERIC) {
				if (parseDateFromVisibleString) {
					String s = getDataFormatter().formatCellValue(cell);
					value = parseDuration(s, pattern);
				} else {
					value = new Date(GenericDateUtil.parseDuration(cell.getNumericCellValue()));
				}
			} else if (cellType == CellType.STRING) {
				String s = getDataFormatter().formatCellValue(cell);
				value = parseDuration(s, pattern);
			}
		}
		return value;
	}

	public boolean readNextRow() {
		int rowIndex = rowStartIndex + currentDatasetNumber;
		if (isCreateStreamingXMLWorkbook()) {
			currentRowIndex = rowIndex;
			currentRow = sheet.getRow(rowIndex);
			if (currentRow != null) {
				rowStartIndex++;
				return true;
			} else {
				return false;
			}
		} else {
			if (rowIndex > maxRowIndex) {
				return false;
			} else {
				currentRowIndex = rowIndex;
				currentRow = sheet.getRow(rowIndex);
				if (currentRow != null) {
					rowStartIndex++;
					return true;
				} else if (stopAtMissingRow) {
					return false;
				} else {
					rowStartIndex++;
					return true;
				}
			}
		}
	}
	
	public void setHeaderName(int columnIndex, String headerName, boolean ignoreMissing) {
		if (headerName != null && headerName.trim().isEmpty() == false) {
			namesFromSchema.put(columnIndex, headerName.trim().toLowerCase());
			ignoreMissingMap.put(columnIndex, ignoreMissing);
		}
	}

	public void configColumnPositions() throws Exception {
		headerRow = sheet.getRow(headerRowIndex);
		int lastCellNum = headerRow.getLastCellNum();
		int firstCellNum = headerRow.getFirstCellNum();
		for (int i = firstCellNum; i <= lastCellNum; i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null) {
				CellType cellType = cell.getCellTypeEnum();
				if (cellType == CellType.STRING) {
					String name = cell.getStringCellValue();
					if (name != null && name.trim().isEmpty() == false) {
						namesFromHeaderRow.put(name.trim().toLowerCase(), i);
					}
				}
			}
		}
		for (Map.Entry<Integer, String> nameFromSchema : namesFromSchema.entrySet()) {
			Boolean ignoreMissing = ignoreMissingMap.get(nameFromSchema.getKey());
			if (ignoreMissing == null) {
				ignoreMissing = false;
			}
			Integer targetIndex = findPosition(nameFromSchema.getValue());
			if (targetIndex != null) {
				columnIndexes.put(nameFromSchema.getKey(), targetIndex);
				individualColumnMappingUsed = true;
			} else if (ignoreMissing) {
				missingColumns.add(nameFromSchema.getKey());
			} else {
				if (findHeaderPosByRegex) {
					throw new Exception("Column with pattern: " + nameFromSchema.getValue() + " does not exists in header!");
				} else {
					throw new Exception("Column with name: " + nameFromSchema.getValue() + " does not exists in header!");
				}
			}
		}
	}
	
	private Integer findPosition(String pattern) {
		if (findHeaderPosByRegex) {
			if (pattern.startsWith("^") == false) {
				pattern = "^" + pattern;
			}
			if (pattern.endsWith("$") == false) {
				pattern = pattern + "$";
			}
			Pattern p = Pattern.compile(pattern, Pattern.CASE_INSENSITIVE);
			for (Map.Entry<String, Integer> entry : namesFromHeaderRow.entrySet()) {
				String header = entry.getKey();
				Integer index = entry.getValue();
				Matcher m = p.matcher(header);
				if (m.find()) {
					return index;
				}
			}
			return null;
		} else {
			return namesFromHeaderRow.get(pattern);
		}
	}
	
	public void useSheet(String sheetName) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		}
		if (sheetName == null || sheetName.trim().isEmpty()) {
			throw new Exception("Name of sheet cannot be null or empty!");
		}
		sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			throw new Exception("Sheet with name:" + targetSheetName + " does not exists!");
		}
		targetSheetName = sheetName;
		currentDatasetNumber = 0;
		sheetLastRowIndex = 0;
		lastValueMap = new HashMap<Integer, Object>();
		maxRowIndex = sheet.getLastRowNum();
	}
	
	public void useSheet(Integer index) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		}
		if (index == null) {
			throw new Exception("Index cannot be null!");
		}
		sheet = workbook.getSheetAt(index);
		if (sheet == null) {
			throw new Exception("Sheet with index:" + index + " does not exists!");
		}
		targetSheetName = sheet.getSheetName();
		currentDatasetNumber = 0;
		sheetLastRowIndex = 0;
		lastValueMap = new HashMap<Integer, Object>();
		maxRowIndex = sheet.getLastRowNum();
	}
	
	public void setDefaultDateFormat(String pattern) {
		if (pattern != null && pattern.trim().length() > 0) {
			defaultDateFormat = new SimpleDateFormat(pattern);
		}
	}
	
	public void setFormatLocale(String locale) {
		setFormatLocale(locale, true);
	}
	
	private Locale createLocale(String locale) {
		int p = locale.indexOf('_');
		String language = locale;
		String country = "";
		if (p > 0) {
			language = locale.substring(0, p);
			country = locale.substring(p);
		}
		return new Locale(language, country);
	}
	
	public void setFormatLocale(String locale, boolean useGrouping) {
		defaultLocale = createLocale(locale);
		defaultNumberFormat = NumberFormat.getInstance(defaultLocale);
		defaultNumberFormat.setMaximumFractionDigits(20);
		defaultNumberFormat.setGroupingUsed(useGrouping);
	}

	public int getHeaderRowIndex() {
		return headerRowIndex;
	}

	public void setHeaderRowIndex(int headerRowIndex) {
		this.headerRowIndex = headerRowIndex;
	}

	public boolean isReturnURLInsteadOfName() {
		return returnURLInsteadOfName;
	}

	public void setReturnURLInsteadOfName(boolean returnURLInsteadOfName) {
		this.returnURLInsteadOfName = returnURLInsteadOfName;
	}

	public boolean isConcatenateLabelUrl() {
		return concatenateLabelUrl;
	}

	public void setConcatenateLabelUrl(boolean concatenateLabelUrl) {
		this.concatenateLabelUrl = concatenateLabelUrl;
	}

	public boolean isFindHeaderPosByRegex() {
		return findHeaderPosByRegex;
	}

	public void setFindHeaderPosByRegex(boolean findHeaderPosByRegex) {
		this.findHeaderPosByRegex = findHeaderPosByRegex;
	}

	public boolean isUseCachedValuesForFailedEvaluation() {
		return useCachedValuesForFailedEvaluations;
	}

	public void setUseCachedValuesForFailedEvaluation(
			boolean useCachedValuesForFailedEvaluations) {
		this.useCachedValuesForFailedEvaluations = useCachedValuesForFailedEvaluations;
	}

	public void setStopAtMissingRow(boolean stopAtMissingRow) {
		this.stopAtMissingRow = stopAtMissingRow;
	}
	
	public boolean rowIsEmpty() {
		if (currentRow == null) {
			return true;
		} else {
			for (Cell cell : currentRow) {
				if (cell != null) {
					if (isCellValueEmpty(cell) == false) {
						return false;
					}
				}
			}
			return true;
		}
	}

	public boolean rowIsEmpty(int ... columns) {
		if (currentRow == null) {
			return true;
		} else {
			for (int i : columns) {
				Cell cell = currentRow.getCell(i);
				if (cell != null) {
					if (isCellValueEmpty(cell) == false) {
						return false;
					}
				}
			}
			return true;
		}
	}

	public int getCurrentRowIndex() {
		return currentRowIndex;
	}

	public boolean isParseDateFromVisibleString() {
		return parseDateFromVisibleString;
	}

	public void setParseDateFromVisibleString(boolean parseDateFromVisibleString) {
		this.parseDateFromVisibleString = parseDateFromVisibleString;
	}

	public boolean isReturnZeroDateAsNull() {
		return returnZeroDateAsNull;
	}

	public void setReturnZeroDateAsNull(boolean returnZeroDateAsNull) {
		this.returnZeroDateAsNull = returnZeroDateAsNull;
	}

	public boolean isLenientDateParsing() {
		return lenientDateParsing;
	}

	public void setLenientDateParsing(Boolean lenientDateParsing) {
		if (lenientDateParsing != null) {
			this.lenientDateParsing = lenientDateParsing;
		}
	}

}