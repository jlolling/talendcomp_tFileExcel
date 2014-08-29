package de.cimt.talendcomp.tfileexcelpoi;

import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class SpreadsheetReferencedCellInput extends SpreadsheetFile {
	
	private Integer currentSheetIndex = null;
	private String currentSheetName = null;
	private Object currentCellValueObject = null;
	private String currentCellValueString = null;
	private Double currentCellValueNumber = null;
	private Boolean currentCellValueBool = null;
	private Date currentCellValueDate = null;
	private String currentCellComment = null;
	private String currentCellCommentAuthor = null;
	private String currentCellFormula = null;
	private String currentCellValueClassName = null;
	private String currentCellBgColor = null;
	private String currentCellFgColor = null;
	private boolean returnURLInsteadOfName = false;
	private Cell currentCell = null;
	private boolean concatenateLabelUrl = false;
	private SimpleDateFormat defaultDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	private NumberFormat numberFormat = NumberFormat.getInstance();

	public boolean readNextCell(String cellRefStr, Object sheetRef, Integer rowIndex, Object columnRef) throws Exception {
		currentSheetIndex = null;
		currentSheetName = null;
		if (sheetRef instanceof Number) {
			currentSheetIndex = ((Number) sheetRef).intValue();
		} else if (sheetRef instanceof String) {
			if (((String) sheetRef).trim().isEmpty() == false) {
				currentSheetName = ((String) sheetRef).trim();
			}
		}
		if (cellRefStr != null && cellRefStr.trim().isEmpty() == false) {
			return readNextCell(cellRefStr);
		} else if (rowIndex != null && columnRef != null) {
			return readNextCell(rowIndex, columnRef);
		} else {
			return false;
		}
	}

	private boolean readNextCell(int rowIndex, Object columnRef) throws Exception {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized!");
		}
		clearCurrentCellValue();
		if (rowIndex < 1) {
			throw new IllegalArgumentException("Row index must >= 1");
		}
		int columnIndex = -1;
		if (columnRef instanceof Number) {
			columnIndex = ((Number) columnRef).intValue();
		} else if (columnRef instanceof String) {
			columnIndex = ExcelCellReference.parseCellColumnName((String) columnRef);
		} else {
			throw new IllegalArgumentException("Cell column refeference must be an none empty String or a number");
		}
		Row row = getSheet().getRow(rowIndex - 1);
		if (row != null) {
			Cell cell = row.getCell(columnIndex);
			if (cell != null) {
				return fetchCurrentCellValue(cell);
			} else {
				return  false;
			}
		} else {
			return false;
		}
	}
	
	private Sheet getSheet() throws Exception {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized!");
		}
		Sheet sheet = null;
		if (currentSheetName != null) {
			sheet = workbook.getSheet(currentSheetName);
			if (sheet == null) {
				throw new Exception("Sheet with name:" + currentSheetName + " does not exists.");
			}
		} else if (currentSheetIndex != null) {
			sheet = workbook.getSheetAt(currentSheetIndex);
			if (sheet == null) {
				throw new Exception("Sheet with index:" + currentSheetIndex + " does not exists.");
			}
		} else {
			throw new Exception("No sheet name or index given!");
		}
		return sheet;
	}
	
	private boolean readNextCell(String cellRefStr) throws Exception {
		if (cellRefStr == null || cellRefStr.trim().isEmpty()) {
			throw new IllegalArgumentException("cellRefStr cannot ne null or empty!");
		}
		CellReference cellRef = new CellReference(cellRefStr.trim());
		String sheetNameFromRef = cellRef.getSheetName();
		if (sheetNameFromRef != null && sheetNameFromRef.trim().isEmpty() == false) {
			currentSheetIndex = null;
			currentSheetName = sheetNameFromRef.trim();
		}
		return readNextCell(cellRef.getRow() + 1, cellRef.getCol());
	}
	
	private void clearCurrentCellValue() {
		currentCellValueClassName = null;
		currentCellValueObject = null;
		currentCellValueString = null;
		currentCellValueNumber = null;
		currentCellValueDate = null;
		currentCellValueBool = null;
		currentCellComment = null;
		currentCellCommentAuthor = null;
		currentCellFormula = null;
		currentCellBgColor = null;
		currentCellFgColor = null;
		currentCell = null;
	}

	private boolean fetchCurrentCellValue(Cell cell) {
		if (cell != null) {
			currentCell = cell;
			currentCellValueString = getStringCellValue(cell);
			Comment comment = cell.getCellComment();
			if (comment != null) {
				currentCellComment = comment.getString().getString();
				currentCellCommentAuthor = comment.getAuthor();
			}
			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				currentCellValueClassName = "Object";
			} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				currentCellValueClassName = "String";
				currentCellValueObject = currentCellValueString;
			} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
				currentCellValueClassName = "Boolean";
				currentCellValueBool = cell.getBooleanCellValue();
				currentCellValueObject = currentCellValueBool;
			} else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
				currentCellValueClassName = "Byte";
				currentCellValueObject = cell.getErrorCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				currentCellValueClassName = "String";
				currentCellFormula = cell.getCellFormula();
				currentCellValueString = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
				currentCellValueObject = currentCellValueString;
			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				if (DateUtil.isCellDateFormatted(cell)) {
					currentCellValueClassName = "java.util.Date";
					currentCellValueDate = cell.getDateCellValue();
					currentCellValueObject = currentCellValueDate;
				} else {
					currentCellValueClassName = "Double";
					currentCellValueNumber = cell.getNumericCellValue();
					currentCellValueObject = currentCellValueNumber;
				}
			}
			currentCellBgColor = getBgColor(cell);
			currentCellFgColor = getFgColor(cell);
			return currentCellValueObject != null;
		} else {
			return false;
		}
	}
	
	private String getBgColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style != null) {
			return getColorString(style.getFillBackgroundColorColor());
		} else {
			return null;
		}
	}

	private String getFgColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style != null) {
			return getColorString(style.getFillForegroundColorColor());
		} else {
			return null;
		}
	}
	
	public static String getColorString(Color color) {
		if (color instanceof HSSFColor) {
			short[] rgb = ((HSSFColor) color).getTriplet();
			if (rgb != null) {
				return (rgb[0] & 0xFF) + ":" + (rgb[1] & 0xFF) + ":" + (rgb[2] & 0xFF);
			} else {
				return null;
			}
		} else if (color instanceof XSSFColor) {
			byte[] rgb = ((XSSFColor) color).getRgb();
			if (rgb != null) {
				return (rgb[0] & 0xFF) + ":" + (rgb[1] & 0xFF) + ":" + (rgb[2] & 0xFF);
			} else {
				return null;
			}
		} else {
			return null;
		}
	}

	public String getCurrentCellComment() {
		return currentCellComment;
	}

	public String getCurrentCellCommentAuthor() {
		return currentCellCommentAuthor;
	}

	public String getCurrentCellFormula() {
		return currentCellFormula;
	}

	public String getCurrentCellValueClassName() {
		return currentCellValueClassName;
	}

	public Object getCurrentCellValueObject() {
		return currentCellValueObject;
	}

	public String getCurrentCellValueString() {
		return currentCellValueString;
	}

	public Double getCurrentCellValueNumber() {
		return currentCellValueNumber;
	}

	public Boolean getCurrentCellValueBool() {
		return currentCellValueBool;
	}

	public Date getCurrentCellValueDate() {
		return currentCellValueDate;
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

	private String getStringCellValue(Cell cell) {
		String value = null;
		if (cell != null) {
			if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				value = getDataFormatter().formatCellValue(cell, getFormulaEvaluator());
			} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
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
			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				if (DateUtil.isCellDateFormatted(cell)) {
					Date d = cell.getDateCellValue();
					value = defaultDateFormat.format(d);
				} else {
					value = numberFormat.format(cell.getNumericCellValue());
				}
			} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
				value = cell.getBooleanCellValue() ? "true" : "false";
			} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				value = null;
			}
		}
		return value;
	}

	public String getCurrentCellBgColor() {
		return currentCellBgColor;
	}

	public String getCurrentCellFgColor() {
		return currentCellFgColor;
	}

	public Cell getCurrentCell() {
		return currentCell;
	}
	
}