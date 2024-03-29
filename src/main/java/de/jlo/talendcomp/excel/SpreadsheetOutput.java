/**
 * Copyright 2018 Jan Lolling jan.lolling@gmail.com
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

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

public class SpreadsheetOutput extends SpreadsheetFile {
	
	private List<Integer> autoSizeColumns = new ArrayList<Integer>();
	private boolean autoSizeAllColumns = false;
	private List<Integer> listColumnsToWriteComment = new ArrayList<Integer>();
	private List<Integer> listColumnsToWriteHyperlink = new ArrayList<Integer>();
	private Drawing<?> drawing = null;
	private boolean groupRowsByColumn = false;
	private Map<Integer, GroupInfo> groupInfoMap = new HashMap<Integer, SpreadsheetOutput.GroupInfo>();
	private boolean writeNullValues = false;
	private Map<Integer, Short> cellFormatMap = new HashMap<Integer, Short>();
	private Map<Integer, CellStyle> columnStyleMap = new HashMap<Integer, CellStyle>();
	private Map<Integer, CellStyle> oddRowColumnStyleMap = new HashMap<Integer, CellStyle>();
	private Map<Integer, CellStyle> evenRowColumnStyleMap = new HashMap<Integer, CellStyle>();
	private boolean reuseExistingStylesFromFirstWrittenRow = false;
	private boolean appendData = false;
	private boolean reuseExistingStylesAlternating = false;
	private boolean reuseFirstRowHeight = false;
	private short firstRowHeight = 800;
	private boolean firstRowIsHeader = false;
	private List<Integer> usedCellColumnIndexes = new ArrayList<Integer>();
	private int commentHeight = 3;
	private int commentWidth = 3;
	private String commentAuthor = null;
	private int dataRowCount = 0;
	private boolean setupCellStylesForAllColumns = false;
	private int highestColumnIndex = 0;
	private boolean writeZeroDateAsNull = true;
	private boolean forbidWritingInProtectedCells = false;
	private int templateRowIndexForStyles = -1;
	
	public void resetCache() {
		if (workbook == null) {
			throw new IllegalStateException("Workbook is not initialized!");
		}
		if (sheet == null) {
			throw new IllegalStateException("Sheet is null. Please take care the setTargetSheetName method is called before.");
		}
		sheetLastRowIndex = sheet.getLastRowNum();
		currentRecordIndex = 0;
		listColumnsToWriteComment.clear();
		listColumnsToWriteHyperlink.clear();
		cellFormatMap.clear();
		autoSizeColumns.clear();
		usedCellColumnIndexes.clear();
		columnStyleMap.clear();
		oddRowColumnStyleMap.clear();
		evenRowColumnStyleMap.clear();
	}
	
	public void freezeAt(int columnIndex, int rowIndex) {
		sheet.createFreezePane(columnIndex, rowIndex, columnIndex, rowIndex);
	}

	public void freezeAt(String columnName, int rowIndex) {
		freezeAt(CellReference.convertColStringToIndex(columnName), rowIndex);
	}

	public void setAutoSizeColumn(int columnIndex) {
		if (autoSizeColumns.contains(Integer.valueOf(columnIndex)) == false) {
			autoSizeColumns.add(columnIndex);
		}
	}
	
	public void writeRow(List<? extends Object> listValues) throws Exception {
		Object[] oneRow = listValues.toArray();
		writeRow(oneRow);
	}

	public void writeColumn(List<? extends Object> listValues) throws Exception {
		Object[] oneRow = listValues.toArray();
		writeColumn(oneRow);
	}

	public void writeColumn(Object[] dataset) throws Exception {
		if (sheet == null) {
			throw new IOException("Sheet is not initialized!");
		}
		int dataColumnIndex = 0;
		for (Object value : dataset) {
			currentRow = getRow(rowStartIndex + dataColumnIndex);
			Cell cell = getCell(currentRow, getCellIndex(currentRecordIndex));
			if (currentRecordIndex == 0) {
				if (autoSizeAllColumns) {
					setAutoSizeColumn(dataColumnIndex);
				}
			}
			writeCellValue(cell, value, dataColumnIndex, currentRecordIndex, false);
			dataColumnIndex++;
		}
		currentRecordIndex++;
	}
	
	/**
	 * shifts the existing rows in row down and creates a new empty row
	 * @param index row index of the new empty inserted row 
	 */
	public void shiftRows(int index) {
		sheet.shiftRows(index, sheet.getLastRowNum(), 1); // move the rows one down
		sheet.createRow(index); // create a new empty row
	}
	
	public void shiftCurrentRow() {
		int start = rowStartIndex + currentRecordIndex;
		int end = sheet.getLastRowNum();
		if (start < end) {
			// only shift if there is a row after the row we are currently writing
			Row shiftedRow = getRow(start);
			sheet.shiftRows(start, end, 1, true, false); // move the rows one down
			Row newRow = getRow(start); // create a new empty row if needed
			newRow.setHeight(shiftedRow.getHeight());
			newRow.setRowStyle(shiftedRow.getRowStyle());
			// copy style from the shifted cells to the new created cells
			for (Cell shiftedCell : shiftedRow) {
				Cell newCell = getCell(newRow, shiftedCell.getColumnIndex());
				newCell.setCellStyle(shiftedCell.getCellStyle());
			}
		}
	}
	
	/**
	 * writes the data in the sheet and creates if necessary a new row.
	 * @param record
	 * @throws Exception
	 */
	public void writeRow(Object[] record) throws Exception {
		dataRowCount = record.length;
		if (sheet == null) {
			throw new IOException("Sheet is not initialized!");
		}
		currentRow = getRow(rowStartIndex + currentRecordIndex);
		if (isFirstRow(currentRecordIndex)) {
			if (appendData) {
				Row pr = sheet.getRow(templateRowIndexForStyles >= 0 ? templateRowIndexForStyles : currentRow.getRowNum());
				if (pr != null) {
					firstRowHeight = pr.getHeight();
					if (reuseFirstRowHeight) {
						currentRow.setHeight(firstRowHeight);
					}
				} else {
					firstRowHeight = currentRow.getHeight();
				}
			} else {
				firstRowHeight = currentRow.getHeight();
			}
		} else if (isDataRow(currentRecordIndex)) {
			if (reuseFirstRowHeight) {
				currentRow.setHeight(firstRowHeight);
			}
		}
		int dataColumnIndex = 0;
		// setup named styles
		for (Object value : record) {
			int cellIndex = getCellIndex(dataColumnIndex);
			Cell cell = getCell(currentRow, cellIndex);
			if (currentRecordIndex == 0) {
				if (autoSizeAllColumns) {
					setAutoSizeColumn(dataColumnIndex);
				}
			}
			writeCellValue(cell, value, dataColumnIndex, currentRecordIndex, false);
			if (groupRowsByColumn) {
				GroupInfo gi = groupInfoMap.get(dataColumnIndex);
				if (gi != null) {
					if (value != null) {
						if (gi.lastValue == null) {
							gi.lastGroupStart = currentRow.getRowNum();
							gi.lastValue = value;
						} else if (value.equals(gi.lastValue) == false) {
							addRowGroup(gi.lastGroupStart, gi.lastRowWithNotNullValue);
							gi.lastGroupStart = currentRow.getRowNum();
							gi.lastValue = value;
						}
						gi.lastRowWithNotNullValue = currentRow.getRowNum();
					}
				}
			}
			if (highestColumnIndex < currentRow.getLastCellNum()) {
				highestColumnIndex = currentRow.getLastCellNum();
			}
			if (usedCellColumnIndexes.contains(cellIndex) == false) {
				usedCellColumnIndexes.add(cellIndex);
			}
			dataColumnIndex++;
		}
		if (setupCellStylesForAllColumns) {
			// must be called as long as currentDatasetNumber points to the current row 
			setupCellStylesForAllUnwrittenColumns();
		}
		currentRecordIndex++;
	}
	
	private void setupCellStylesForAllUnwrittenColumns() {
		// setup style from all other columns in the row
		for (int ci = 0; ci < highestColumnIndex; ci++) {
			if (usedCellColumnIndexes.contains(ci) == false) {
				Cell cell = currentRow.getCell(ci);
				if (cell == null) {
					cell = currentRow.createCell(ci);
				}
				setupStyle(cell, currentRecordIndex);
			}
		}
	}
	
	private static class GroupInfo {
		
		Object lastValue = null;
		int lastRowWithNotNullValue = 0;
		int lastGroupStart = 0;
		
	}
	
	public void closeLastGroup() {
		if (groupRowsByColumn) {
			for (GroupInfo gi : groupInfoMap.values()) {
				if (gi.lastGroupStart < gi.lastRowWithNotNullValue - 1) {
					addRowGroup(gi.lastGroupStart, gi.lastRowWithNotNullValue);
				}
			}
		}
	}
	
	public boolean writeNamedCellValue(String cellName, Object value) throws Exception {
		Cell cell = getNamedCell(cellName);
		if (cell != null) {
			writeCellValue(cell, value, -1, -1, true);
			return true;
		} else {
			return false;
		}
	}

	private void setupReferencedSheet(String cellRefStr, Object sheetRef) throws Exception {
		if (sheetRef instanceof String) {
			sheet = workbook.getSheet((String) sheetRef);
			if (sheet == null) {
				sheet = workbook.createSheet((String) sheetRef);
			}
		} else if (sheetRef instanceof Number) {
			int sheetIndex = ((Number) sheetRef).intValue();
			if (workbook.getNumberOfSheets() <= sheetIndex) {
				throw new Exception("Sheet with index: " + ((Number) sheetRef).intValue() + " does not exists and can only be created if a sheet name is provided");
			}
			sheet = workbook.getSheetAt(sheetIndex);
			if (sheet == null) {
				throw new Exception("Sheet with index: " + ((Number) sheetRef).intValue() + " does not exists and can only be created if a sheet name is provided");
			}
		} else if (cellRefStr != null && cellRefStr.trim().isEmpty() == false) {
			CellReference cellRef = new CellReference(cellRefStr.trim());
			String sheetNameFromRef = cellRef.getSheetName();
			if (sheetNameFromRef != null && sheetNameFromRef.trim().isEmpty() == false) {
				sheet = workbook.getSheet(sheetNameFromRef);
				if (sheet == null) {
					sheet = workbook.createSheet(sheetNameFromRef);
				}
			}
		}
	}
	
	public void writeReferencedCellValue(String cellRefStr, Object sheetRef, Integer rowIndex, Object columnRef, Object value, String comment, Boolean doNotSetCellStyle) throws Exception {
		setupReferencedSheet(cellRefStr, sheetRef);
		if (cellRefStr != null && cellRefStr.trim().isEmpty() == false) {
			CellReference cellRef = new CellReference(cellRefStr.trim());
			// cellRef is for row 0-based
			writeReferencedCellValue(cellRef.getRow(), cellRef.getCol(), value, comment, doNotSetCellStyle);
		} else {
			if (rowIndex == null || rowIndex < 0) {
				throw new Exception("row index cannot be null or lower 0");
			}
			if (columnRef == null) {
				throw new Exception("column index cannot be null");
			}
			writeReferencedCellValue(rowIndex - 1, columnRef, value, comment, doNotSetCellStyle);
		}
	}
	
	private void writeReferencedCellValue(Integer rowIndex, Object column, Object value, String comment, Boolean doNotSetCellStyle) throws Exception {
		if (rowIndex == null || rowIndex < 0) {
			throw new Exception("row index cannot be null or lower 0");
		}
		if (column == null) {
			throw new Exception("column index cannot be null");
		}
		int columnIndex = 0;
		if (column instanceof String) {
			columnIndex = CellReference.convertColStringToIndex((String) column);
		} else if (column instanceof Number) {
			columnIndex = ((Number) column).intValue();
		} else {
			throw new Exception("The value " + column + " in parameter column cannot be used as column index. Value must be a String (excel column name) or a number (column index 0-based)");
		}
		if (columnIndex < 0) {
			throw new Exception("column index must be 0 or greater");
		}
		if (sheet == null) {
			throw new IOException("Sheet is not initialized!");
		}
		Row row = getRow(rowIndex);
		Cell cell = getCell(row, columnIndex);
		writeCellValue(cell, value, columnIndex, rowIndex, (doNotSetCellStyle != null ? doNotSetCellStyle.booleanValue() : false));
		if (comment != null && comment.isEmpty() == false) {
			setCellComment(cell, comment);
		}
	}
	
	private Drawing<?> getDrawing() {
		if (drawing == null) {
			drawing = sheet.createDrawingPatriarch();
		}
		return drawing;
	}
	
	private void setCellComment(Cell cell, String comment) {
		if (comment == null || comment.trim().isEmpty()) {
			cell.removeCellComment();
		} else {
			Comment c = cell.getCellComment();
			if (c == null) {
				ClientAnchor anchor = creationHelper.createClientAnchor();
				anchor.setRow1(cell.getRowIndex());
				anchor.setRow2(cell.getRowIndex() + commentHeight);
				anchor.setCol1(cell.getColumnIndex());
				anchor.setCol2(cell.getColumnIndex() + commentWidth);
				anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
				c = getDrawing().createCellComment(anchor);
				c.setVisible(false);
				if (commentAuthor != null) {
					c.setAuthor(commentAuthor);
				}
				cell.setCellComment(c);
			}
			RichTextString rts = creationHelper.createRichTextString(comment);
			c.setString(rts);
		}
	}

	private void setCellHyperLink(Cell cell, String url) {
		if (url.contains("://")) {
			Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
			link.setAddress(url);
			cell.setHyperlink(link);
		} else if (url.startsWith("mailto:")) {
			Hyperlink link = creationHelper.createHyperlink(HyperlinkType.EMAIL);
			link.setAddress(url);
			cell.setHyperlink(link);
		} else {
			Hyperlink link = creationHelper.createHyperlink(HyperlinkType.FILE);
			link.setAddress(url);
			cell.setHyperlink(link);
		}
		if (cell.getCellType() == CellType.BLANK) {
			cell.setCellValue(url);
		}
	}
	
	private boolean isCellProtected(Cell cell) {
		if (cell != null) {
			return sheet.getProtect() && cell.getCellStyle().getLocked();
		}
		return false;
	}
	
	private void writeCellValue(Cell cell, Object value, int dataColumnIndex, int dataRowIndex, boolean doNotSetCellStyle) throws Exception {
		if (forbidWritingInProtectedCells) {
			if (isCellProtected(cell)) {
				throw new Exception("The cell: " + new CellReference(cell).formatAsString() + " is locked and option forbid writing in protected cells is enabled.");
			}
		}
		if (value instanceof String) {
			String s = (String) value;
			boolean isPlainValue = true;
			if (isToWriteAsComment(dataColumnIndex)) {
				// if this schema data column is dedicated as comment 
				isPlainValue = false;
				if (firstRowIsHeader == false || dataRowIndex > 0) {
					// avoid set comment for the header line
					setCellComment(cell, s);
				}
			}
			if (isToWriteAsLink(dataColumnIndex)) {
				// if this schema data column is dedicated as hyper link
				if (firstRowIsHeader == false || dataRowIndex > 0) {
					// avoid set hyper-links for the header line
					setCellHyperLink(cell, s);
					isPlainValue = false;
				}
			}
			if (isPlainValue) {
				if (s.startsWith("=")) {
					int rowNum = cell.getRow().getRowNum();
					cell.setCellFormula(getFormular(s, rowNum));
				} else {
					cell.setCellValue(s);
					if (firstRowIsHeader == false || dataRowIndex > 0) {
						// do not mislead the cell format if we are in the header line
						cellFormatMap.put(cell.getColumnIndex(), (short) org.apache.poi.ss.usermodel.BuiltinFormats.getBuiltinFormat("TEXT"));
					}
				}
			}
		} else if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Long) {
			cell.setCellValue((Long) value);
		} else if (value instanceof BigInteger) {
			cell.setCellValue(((BigInteger) value).longValue());
		} else if (value instanceof BigDecimal) {
			cell.setCellValue(((BigDecimal) value).doubleValue());
		} else if (value instanceof Double) {
			cell.setCellValue((Double) value);
		} else if (value instanceof Float) {
			cell.setCellValue((Float) value);
		} else if (value instanceof Short) {
			cell.setCellValue((Short) value);
		} else if (value instanceof Number) {
			cell.setCellValue(Double.valueOf(((Number) value).doubleValue()));
		} else if (value instanceof Date) {
			if (writeZeroDateAsNull && GenericDateUtil.isZeroDate((Date) value)) {
				cell.setBlank();
			} else {
				cell.setCellValue((Date) value);
			}
		} else if (value != null) {
			cell.setCellValue(value.toString());
		} else if (writeNullValues && value == null) {
			cell.setBlank();
		}
		if (isDataRow(dataRowIndex) && doNotSetCellStyle == false) {
			setupStyle(cell, dataRowIndex);
		}
	}
	
	public void setupColumnSize() {
		for (Integer ci : autoSizeColumns) {
			sheet.autoSizeColumn(getCellIndex(ci.intValue()));
		}
	}
	
	public boolean isAutoSizeAllColumns() {
		return autoSizeAllColumns;
	}

	public void setAutoSizeAllColumns(boolean autoSizeAllColumns) {
		this.autoSizeAllColumns = autoSizeAllColumns;
	}

	public Sheet createCopy(Sheet sourceSheet, String targetSheetName) throws Exception {
		if (sourceSheet == null) {
			throw new IllegalArgumentException("Create copy of sheet failed: Source sheet cannot be null");
		}
		int sourceSheetIndex = workbook.getSheetIndex(sourceSheet);
		return createCopy(sourceSheetIndex, targetSheetName);
	}

	public Sheet createCopy(String sourceSheetName, String targetSheetName) throws Exception {
		if (sourceSheetName == null || sourceSheetName.trim().isEmpty()) {
			throw new IllegalArgumentException("Create copy of sheet failed: Source sheet name cannot be null or empty");
		}
		int sourceSheetIndex = workbook.getSheetIndex(sourceSheetName);
		if (sourceSheetIndex < 0) {
			throw new Exception("Create a copy of sheet: " + sourceSheetName + " failed. This sheet does not exist in the current workbook.");
		}
		return createCopy(sourceSheetIndex, targetSheetName);
	}

	public Sheet createCopy(Integer sourceSheetIndex, String targetSheetName) throws Exception {
		if (sourceSheetIndex == null || sourceSheetIndex < 0) {
			throw new IllegalArgumentException("Create copy of sheet failed. Source sheet index: " + sourceSheetIndex + " must be 0 or greater.");
		}
		if (targetSheetName == null || targetSheetName.trim().isEmpty()) {
			throw new IllegalArgumentException("Create copy of sheet with index: " + sourceSheetIndex + " failed: Target sheet name cannot be null or empty");
		}
		try {
			Sheet newSheet = workbook.cloneSheet(sourceSheetIndex);
			this.targetSheetName = ensureCorrectExcelSheetName(targetSheetName);
			workbook.setSheetName(workbook.getSheetIndex(newSheet), targetSheetName);
			return newSheet;
		} catch (Throwable t) {
			if (workbook instanceof SXSSFWorkbook) {
				throw new Exception("Copying a sheet cannot work in a workbook which is not fully loaded because of the memory saving mode. Uncheck Memory saving mode in tFileExcelWorkbookOpen!", t);
			} else {
				throw new Exception("createCopy from source failed:" + t.getMessage(), t);
			}
		}
	}

	@Override
	public int getLastRowNum() {
		if (currentRow != null) {
			return currentRow.getRowNum();
		} else {
			return -1;
		}
	}
	
	public void deleteFollowingRows() {
    	if (workbook instanceof SXSSFWorkbook) {
			warn("Cannot delete following rows in the memory the saving mode (use of the streaming-workbook).");
    	} else {
			int rowIndex = firstRowIsHeader ? rowStartIndex + 1 : rowStartIndex;
			if (currentRow != null) {
				rowIndex = currentRow.getRowNum() + 1;
			}
			int lastSheetRowIndex = sheet.getLastRowNum();
			for ( ; rowIndex <= lastSheetRowIndex; lastSheetRowIndex--) {
				Row row = sheet.getRow(lastSheetRowIndex);
				if (row != null) {
					sheet.removeRow(row);
				}
			}
    	}
	}
	
	private boolean isToWriteAsComment(int columnIndex) {
		for (Integer cc : listColumnsToWriteComment) {
			if (cc.intValue() == columnIndex) {
				return true;
			}
		}
		return false;
	}
	
	private boolean isToWriteAsLink(int columnIndex) {
		for (Integer cc : listColumnsToWriteHyperlink) {
			if (cc.intValue() == columnIndex) {
				return true;
			}
		}
		return false;
	}

	public void setColumnValueAsComment(Integer dataColumnIndex) {
		if (dataColumnIndex != null) {
			listColumnsToWriteComment.add(dataColumnIndex);
		}
	}
	
	public void setColumnValueAsLink(Integer dataColumnIndex) {
		if (dataColumnIndex != null) {
			listColumnsToWriteHyperlink.add(dataColumnIndex);
		}
	}

	public void addColumnGroup(String columnDesc) {
		if (columnDesc != null && columnDesc.trim().isEmpty() == false) {
			String[] groups = columnDesc.split(",");
			if (groups != null) {
				for (String group : groups) {
					String[] cols = group.split("-");
					if (cols.length == 2) {
						addColumnGroup(cols[0].trim(), cols[1].trim());
					}
				}
			}
		}
	}
	
	public void addColumnGroup(Object fromColumn, Object toColumn) {
		int fromColumnIndex = -1;
		if (fromColumn instanceof Number) {
			fromColumnIndex = ((Number) fromColumn).intValue();
		} else if (fromColumn instanceof String) {
			fromColumnIndex = CellReference.convertColStringToIndex((String) fromColumn);
		}
		int toColumnIndex = 0;
		if (toColumn instanceof Number) {
			toColumnIndex = ((Number) toColumn).intValue();
		} else if (toColumn instanceof String) {
			toColumnIndex = CellReference.convertColStringToIndex((String) toColumn);
		}
		if (fromColumnIndex >= 0 && fromColumnIndex < toColumnIndex - 1) {
			sheet.groupColumn(fromColumnIndex, toColumnIndex);
		}
	}
	
	public void addRowGroup(int fromRow, int toRow) {
		if (fromRow < toRow - 1) {
			sheet.groupRow(fromRow, toRow - 1);
		}
	}

	public void groupRowsByColumn(String columnName) {
		if (columnName != null && columnName.trim().isEmpty() == false) {
			groupRowsByColumn = true;
			int columnIndex = CellReference.convertColStringToIndex(columnName);
			if (groupInfoMap.get(columnIndex) == null) {
				GroupInfo gi = new GroupInfo();
				groupInfoMap.put(columnIndex, gi);
			}
		}
	}
	
	public void groupRowsByColumn(Integer ... columnIndexes) {
		if (columnIndexes != null) {
			for (Integer columnIndex : columnIndexes) {
				groupRowsByColumn = true;
				if (groupInfoMap.get(columnIndex) == null) {
					GroupInfo gi = new GroupInfo();
					groupInfoMap.put(columnIndex, gi);
				}
			}
		}
	}

	/**
	 * set the number format for data row column
	 * @param schemaColumnIndex index of column in data (cell index can differ: see setColumnMapping)
	 * @param pattern #,##0.00 means thousand delimiter and precision 2
	 */
	public void setDataFormat(int schemaColumnIndex, String pattern) {
		if (pattern != null && pattern.trim().isEmpty() == false) {
			pattern = pattern.replace("yy", "YY").replace("dd", "DD"); // converts Java to Excel format
			short formatIndex = format.getFormat(pattern);
			Integer cellColumnPos = columnIndexes.get(schemaColumnIndex);
			if (cellColumnPos != null) {
				cellFormatMap.put(cellColumnPos, formatIndex);
			} else {
				cellFormatMap.put(schemaColumnIndex, formatIndex);
			}
		}
	}
	
	public void setNumberPrecision(int schemaColumnIndex, int precision) {
		short formatIndex = format.getFormat(createPrecisionPattern(precision));
		Integer cellColumnPos = columnIndexes.get(schemaColumnIndex);
		if (cellColumnPos != null) {
			cellFormatMap.put(cellColumnPos, formatIndex);
		} else {
			cellFormatMap.put(schemaColumnIndex, formatIndex);
		}
	}
	
	private String createPrecisionPattern(int precision) {
		StringBuilder pattern = new StringBuilder("#,##0");
		for (int i = 0; i < precision; i++) {
			if (i == 0) {
				pattern.append(".");
			}
			pattern.append("0");
		}
		return pattern.toString();
	}

	public boolean isWriteNullValues() {
		return writeNullValues;
	}

	public void setWriteNullValues(boolean writeNullValues) {
		this.writeNullValues = writeNullValues;
	}
	
	private boolean isFirstRow(int row) {
		if (firstRowIsHeader) {
			return row == 1;
		} else {
			return row == 0;
		}
	}
	
	private boolean isDataRow(int row) {
		if (firstRowIsHeader) {
			return row > 0;
		} else if (row >= 0) {
			return true;
		} else {
			return false;
		}
	}

	private boolean isSecondRow(int row) {
		if (firstRowIsHeader) {
			return row == 2;
		} else {
			return row == 1;
		}
	}

	private boolean isEvenDataRow(int row) {
		if (firstRowIsHeader) {
			return row % 2 == 0;
		} else {
			return row % 2 > 0;
		}
	}

	private void setupStyle(Cell cell, int dataRowIndex) {
		// cell has its own style and not the default style
		if (reuseExistingStylesFromFirstWrittenRow) {
			if (appendData == false) {
				// we have to reuse the existing style
				// from the first written row
				if (reuseExistingStylesAlternating) {
					CellStyle style = cell.getCellStyle();
					// we have to reuse the style from the even/odd row
					if (isFirstRow(dataRowIndex)) {
						// we are in the first row, memorize the style
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							oddRowColumnStyleMap.put(cell.getColumnIndex(), style);
						}
					} else if (isSecondRow(dataRowIndex)) {
						// we are in the first row, memorize the style
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							evenRowColumnStyleMap.put(cell.getColumnIndex(), style);
						}
					} else if (isEvenDataRow(dataRowIndex)) {
						// reference to the previously memorized style for even rows
						CellStyle s = evenRowColumnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					} else {
						// reference to the previously memorized style for even rows
						CellStyle s = oddRowColumnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					}
				} else {
					// here we go if we append data
					// we have to take the styles from the previous rows
					CellStyle style = cell.getCellStyle();
					// we take the style from the last row
					if (isFirstRow(dataRowIndex)) {
						// memorize the style for reuse in all other rows
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							columnStyleMap.put(cell.getColumnIndex(), style);
						}
					} else {
						// set the style from the previous row
						CellStyle s = columnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					}
				}
			} else {
				// append mode
				// we have to get the styles from the previous rows
				if (reuseExistingStylesAlternating) {
					CellStyle style = cell.getCellStyle();
					// we have to reuse the style from the even/odd row
					if (isFirstRow(dataRowIndex)) {
						// get the cell from the pre-previous row
						if (cell.getRowIndex() > 1) {
							// we have 2 rows above from which we can take the styles
							Row ppr = sheet.getRow(templateRowIndexForStyles >= 0 ? templateRowIndexForStyles : cell.getRowIndex() - 2);
							if (ppr != null) {
								Cell pprc = ppr.getCell(cell.getColumnIndex());
								if (pprc != null) {
									style = pprc.getCellStyle();
								}
							}
						}
						// we are in the first row, memorize the style
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							oddRowColumnStyleMap.put(cell.getColumnIndex(), style);
							cell.setCellStyle(style);
						}
					} else if (isSecondRow(dataRowIndex)) {
						// get the cell from the pre-previous row
						if (cell.getRowIndex() > 1) {
							// we have 2 rows above from which we can take the styles
							Row ppr = sheet.getRow(templateRowIndexForStyles >= 0 ? templateRowIndexForStyles + 1 : cell.getRowIndex() - 2);
							if (ppr != null) {
								Cell pprc = ppr.getCell(cell.getColumnIndex());
								if (pprc != null) {
									style = pprc.getCellStyle();
								}
							}
						}
						// we are in the first row, memorize the style
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							evenRowColumnStyleMap.put(cell.getColumnIndex(), style);
							cell.setCellStyle(style);
						}
					} else if (isEvenDataRow(dataRowIndex)) {
						// reference to the previously memorized style for even rows
						CellStyle s = evenRowColumnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					} else {
						// reference to the previously memorized style for even rows
						CellStyle s = oddRowColumnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					}
				} else {
					// we take the style from the last row
					if (isFirstRow(dataRowIndex)) {
						CellStyle style = cell.getCellStyle();
						// get the cell from the previous row
						if (cell.getRowIndex() > 0) {
							// we have 1 rows above from which we can take the styles
							Row pr = sheet.getRow(templateRowIndexForStyles >= 0 ? templateRowIndexForStyles : cell.getRowIndex() - 1);
							if (pr != null) {
								Cell prc = pr.getCell(cell.getColumnIndex());
								if (prc != null) {
									style = prc.getCellStyle();
								}
							}
						}
						// memorize the style for reuse in all other rows
						if (style.getIndex() > 0) {
							// only if the cell does not use the default style
							columnStyleMap.put(cell.getColumnIndex(), style);
							cell.setCellStyle(style);
						}
					} else {
						// set the style from the previous row
						CellStyle s = columnStyleMap.get(cell.getColumnIndex());
						if (s != null) {
							cell.setCellStyle(s);
						}
					}
				}
			}
		} else { // reuseExistingStylesFromFirstWrittenRow == false
			Short formatIndex = cellFormatMap.get(cell.getColumnIndex());
			if (formatIndex != null) {
				CellStyle style = cell.getCellStyle();
				if ((style.getIndex() == 0) || (style.getDataFormat() != formatIndex.shortValue())) {
					// this is the default style or the current format differs from the given format
					// we need our own style for this 
					CellStyle newstyle = columnStyleMap.get(cell.getColumnIndex());
					if (newstyle == null) {
						newstyle = createCellStyle(style);
						newstyle.setDataFormat(formatIndex.shortValue());
						columnStyleMap.put(cell.getColumnIndex(), newstyle);
					}
					cell.setCellStyle(newstyle);
				}
			}
		}
	}
	
	private CellStyle createCellStyle(CellStyle template) {
		CellStyle newstyle = workbook.createCellStyle();
		if (template != null) {
			newstyle.cloneStyleFrom(template);
		}
		return newstyle;
	}

	public boolean isReuseExistingStylesFromFirstWrittenRow() {
		return reuseExistingStylesFromFirstWrittenRow;
	}

	public void setReuseExistingStylesFromFirstWrittenRow(boolean reuseExistingStyles) {
		this.reuseExistingStylesFromFirstWrittenRow = reuseExistingStyles;
	}

	public void setAppend(boolean reuseExistingStyles) {
		this.appendData = reuseExistingStyles;
	}

	public boolean isReuseExistingStylesAlternating() {
		return reuseExistingStylesAlternating;
	}

	public void setReuseExistingStylesAlternating(
			boolean reuseExistingStylesAlternating) {
		this.reuseExistingStylesAlternating = reuseExistingStylesAlternating;
	}

	public boolean isFirstRowIsHeader() {
		return firstRowIsHeader;
	}

	public void setFirstRowIsHeader(boolean firstRowIsHeader) {
		this.firstRowIsHeader = firstRowIsHeader;
	}
	
	private ConditionalFormatting currentCf = null;
	private int currentCfIndex = -1;
	private int maxRuleChunkSize = 3; // not higher than 3 because of Excel 2007

	
    private void find(SheetConditionalFormatting scf, int row, int col) {
    	currentCf = null;
    	currentCfIndex = -1;
    	ConditionalFormatting cf = null;
		int numCF = scf.getNumConditionalFormattings();
		for (int i = 0; i < numCF; i++) {
			cf = scf.getConditionalFormattingAt(i);
			CellRangeAddress[] crArray = cf.getFormattingRanges();
			for (CellRangeAddress cra : crArray) {
				if (cra.isInRange(row, col)) {
					if (cra.isFullRowRange() == false) {
						currentCf = cf;
						currentCfIndex = i;
						break;
					} else {
						currentCf = null;
					}
				}
			}
		}
    }
    
	public void extendCellRangeForTable() throws Exception {
		info("Extending cell range for tables...");
		if (sheet instanceof XSSFSheet) {
			int firstDataRowIndex = firstRowIsHeader ? rowStartIndex + 1 : rowStartIndex;
			// check if we are in the append mode and if we are really appending data to existing records
			if (appendData) {
				if (rowStartIndex > 1) {
					firstDataRowIndex--;
				}
			}
			List<XSSFTable> listTables =  ((XSSFSheet) sheet).getTables();
			if (individualColumnMappingUsed) {
				for (Integer col : columnIndexes.values()) {
					// walk through all written columns and ...
					for (int i = listTables.size() - 1; i >= 0; i--) {
						// ... through all tables
						XSSFTable table = listTables.get(i);
						// check if the table is written
						if (extendTable(table, firstDataRowIndex, col.intValue(), getLastRowNum())) {
							// if extended, remove it from the list
							listTables.remove(table);
						}
					}
				}
			} else { 
				for (int col = columnStartIndex; col < (columnStartIndex + dataRowCount); col++) {
					// walk through all written columns and ...
					for (int i = listTables.size() - 1; i >= 0; i--) {
						// ... through all tables
						XSSFTable table = listTables.get(i);
						// check if the table is written
						if (extendTable(table, firstDataRowIndex, col, getLastRowNum())) {
							// if extended, remove it from the list
							listTables.remove(table);
						}
					}
				}
			}
		} else if (workbook instanceof SXSSFWorkbook) {
			warn("Cannot extend cell ranges for tables in the memory saving mode (use of the streaming-workbook).");
		}
	}
	
	private boolean extendTable(XSSFTable table, int firstRow, int firstCol, int lastRow) throws Exception {
		try {
			AreaReference currentRef = null;
			SpreadsheetVersion sv = null;
			if (currentType == SpreadsheetTyp.XLS) {
				sv = SpreadsheetVersion.EXCEL97;
				currentRef = new AreaReference(table.getCTTable().getRef(), sv);
			} else {
				sv = SpreadsheetVersion.EXCEL2007;
				currentRef = new AreaReference(table.getCTTable().getRef(), sv);
			}
			CellReference topLeft = currentRef.getFirstCell();
			CellReference buttomRight = currentRef.getLastCell();
			if (topLeft.getRow() <= firstRow && buttomRight.getRow() >= firstRow && topLeft.getCol() <= firstCol && buttomRight.getCol() >= firstCol) {
				// this table is within our write area, we have to expand it
				AreaReference newRef = new AreaReference(
						topLeft, // left top including the header line
						new CellReference(lastRow, buttomRight.getCol()), // bottom right
						sv); // Excel Version
				table.getCTTable().setRef(newRef.formatAsString());
				return true;
			} else {
				return false;
			}
		} catch (Exception t) {
        	if (workbook instanceof SXSSFWorkbook) {
        		throw new Exception("Extending table ranges cannot work in a workbook which is not fully loaded because of the memory saving mode. Uncheck Memory saving mode in tFileExcelWorkbookOpen!", t);
        	} else {
        		throw t;
        	}
		}
	}
	
    public void extendCellRangesForConditionalFormattings() throws Exception {
    	if (workbook instanceof SXSSFWorkbook) {
			warn("Cannot extend cell ranges for conditional formats in the memory the saving mode (use of the streaming-workbook).");
    	} else {
    		int firstDataRowIndex = firstRowIsHeader ? rowStartIndex + 1 : rowStartIndex;
			// check if we are in the append mode and if we are really appending data to existing records
			if (appendData) {
				if (rowStartIndex > 1) {
					firstDataRowIndex--;
				}
			}
        	info("Extending cell ranges for conditional formats. Use formats from row: " + firstDataRowIndex);
        	if (getLastRowNum() > 0 && getLastRowNum() > firstDataRowIndex) {
        		SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        		if (debug) {
        			debug("#### Conditional formattings before:");
        			debug(logoutSheetConditionalFormatting(scf));
        		}
        		ConditionalFormatting lastCf = null;
            	for (Integer cellColumnIndex : usedCellColumnIndexes) {
            		if (debug) {
            			debug("extendCellRangesForConditionalFormattings: check format for cell index=" + cellColumnIndex);
            		}
            		find(scf, firstDataRowIndex, cellColumnIndex); // currentCf and currentCfIndex will be set here
            		if (currentCf != null && currentCf != lastCf) {
                		if (debug) {
                			debug("extendCellRangesForConditionalFormattings: found format for cell index=" + cellColumnIndex);
                		}
                		lastCf = currentCf;
                		CellRangeAddress[] ranges = currentCf.getFormattingRanges();
                		for (int i = 0; i < ranges.length; i++) {
                			CellRangeAddress address = ranges[i];
                			CellRangeAddress extendedAddress = new CellRangeAddress(address.getFirstRow(), getLastRowNum(), address.getFirstColumn(), address.getLastColumn());
                			ranges[i] = extendedAddress;
                		}
                		if (debug) {
                			debug("extendCellRangesForConditionalFormattings: extend ranges to=" + firstDataRowIndex + ":" + getLastRowNum() + " -> " + getRangesAsString(ranges));
                		}
            			int numRulesTotal = currentCf.getNumberOfRules();
            			if (numRulesTotal > 0) {
            				int chunks = numRulesTotal / maxRuleChunkSize;
            				int restChunkSize = numRulesTotal % maxRuleChunkSize;
            				int currentSize = 0;
            				for (int c = 0; c <= chunks; c++) {
            					if (c < chunks) {
            						// all not-last chunks have the max chunk size
            						currentSize = maxRuleChunkSize;
            					} else {
            						// the last chunk contains the rest
            						currentSize = restChunkSize;
            					}
            					if (currentSize > 0) {
                					ConditionalFormattingRule[] rules = new ConditionalFormattingRule[currentSize];
                        			for (int i = 0; i < currentSize; i++) {
                        				int ruleIndex = i + (maxRuleChunkSize * c); // current pointer within a chunk + chunk offset
                        				rules[i] = currentCf.getRule(ruleIndex);
                                		if (debug) {
                                			debug("extendCellRangesForConditionalFormattings: add ranges: " + getRangesAsString(ranges) + " rule #" + ruleIndex + " =" + describeRule(rules[i]));
                                		}
                        			}
                    				scf.addConditionalFormatting(ranges, rules);
            					}
            				}
                    		if (debug) {
                    			debug("extendCellRangesForConditionalFormattings: remove template format at index:" + currentCfIndex);
                    		}
            				scf.removeConditionalFormatting(currentCfIndex);
            			}
            		}
            	}
        		if (debug) {
        			debug("#### Conditional formattings after:");
        			debug(logoutSheetConditionalFormatting(scf));
        		}
        	}
    	}
    }
    
    private String getRangesAsString(CellRangeAddress[] ranges) {
    	if (ranges != null && ranges.length > 0) {
    		StringBuilder sb = new StringBuilder();
    		for (int i = 0; i < ranges.length; i++) {
    			if (i > 0) {
    				sb.append(";");
    			}
    			sb.append("[");
    			sb.append(ranges[i].formatAsString());
    			sb.append("]");
    		}
    		return sb.toString();
    	}
    	return "";
    }
    
    private String logoutSheetConditionalFormatting(SheetConditionalFormatting scf) {
    	StringBuilder sb = new StringBuilder();
    	int countCf = scf.getNumConditionalFormattings();
    	sb.append("\n");
    	for (int f = 0; f < countCf; f++) {
    		sb.append(logoutConditionalFormat(scf.getConditionalFormattingAt(f)));
    		sb.append("\n");
    	}
    	return sb.toString();
    }
    
    private String logoutConditionalFormat(ConditionalFormatting cf) {
    	StringBuilder sb = new StringBuilder();
    	sb.append("Conditional Format:\n  Ranges:\n    ");
    	CellRangeAddress[] ranges = cf.getFormattingRanges();
    	if (ranges != null) {
    		for (int r = 0; r < ranges.length; r++) {
    			if (r > 0) {
    				sb.append("\n    ");
    			}
    			sb.append(ranges[r].formatAsString());
    		}
    	}
    	sb.append("\n  Rules:\n    ");
    	int nbRules = cf.getNumberOfRules();
    	for (int r = 0; r < nbRules; r++) {
    		if (r > 0) {
    			sb.append("\n    ");
    		}
			sb.append("#" + r + ":");
    		sb.append(describeRule(cf.getRule(r)));
    	}
    	return sb.toString();
    }
        
    private static String describeRuleComparisonOperator(ConditionalFormattingRule rule) {
    	StringBuilder sb = new StringBuilder();
    	sb.append(" comparison:");
    	switch (rule.getComparisonOperation()) {
        case ComparisonOperator.LT: 
        	sb.append(rule.getFormula1());
        	sb.append(" < ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.LE: 
        	sb.append(rule.getFormula1());
        	sb.append(" <= ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.GT: 
        	sb.append(rule.getFormula1());
        	sb.append(" > ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.GE: 
        	sb.append(rule.getFormula1());
        	sb.append(" >= ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.EQUAL: 
        	sb.append(rule.getFormula1());
        	sb.append(" = ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.NOT_EQUAL: 
        	sb.append(rule.getFormula1());
        	sb.append(" != ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.BETWEEN: 
        	sb.append(rule.getFormula1());
        	sb.append(" between ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.NOT_BETWEEN: 
        	sb.append(rule.getFormula1());
        	sb.append(" not between ");
        	sb.append(rule.getFormula2());
        	break;
        case ComparisonOperator.NO_COMPARISON: 
    		sb.append(" none ");
        	break;
    	}
    	return sb.toString();
    }
    
    private static String describeRule(ConditionalFormattingRule rule) {
    	StringBuilder sb = new StringBuilder();
		sb.append("condition:");
		ConditionType ct = rule.getConditionType();
    	if (ct.equals(ConditionType.CELL_VALUE_IS)) {
    		sb.append(" cell value is: ");
    		sb.append(describeRuleComparisonOperator(rule));
    	} else if (ct.equals(ConditionType.FORMULA)) {
    		sb.append(" formula: ");
        	sb.append(rule.getFormula1());
    	} else if (ct.equals(ConditionType.FILTER)) {
    		sb.append(" filter: ");
    		sb.append(describeRuleComparisonOperator(rule));
    	} else if (ct.equals(ConditionType.ICON_SET)) {
    		sb.append(" icon set: ");
    		sb.append(rule.getMultiStateFormatting());
    	} else if (ct.equals(ConditionType.COLOR_SCALE)) {
    		sb.append(" color-scale: ");
    		sb.append(rule.getColorScaleFormatting());
    	} else if (ct.equals(ConditionType.DATA_BAR)) {
    		sb.append(" data-bar: ");
    		sb.append(rule.getDataBarFormatting());
    	} else {
        	sb.append(" type=" + rule.getConditionType());
    	}
    	sb.append(" formattings:");
    	if (rule.getBorderFormatting() != null) {
    		sb.append(" [has border formats]");
    	}
    	if (rule.getFontFormatting() != null) {
    		sb.append(" [has font formattings]");
    	}
    	if (rule.getPatternFormatting() != null) {
    		sb.append(" [has pattern formattings]");
    	}
    	return sb.toString();
    }

	public boolean isReuseFirstRowHeight() {
		return reuseFirstRowHeight;
	}

	public void setReuseFirstRowHeight(boolean reuseFirstRowHeight) {
		this.reuseFirstRowHeight = reuseFirstRowHeight;
	}
	
	public Sheet getSheet() {
		if (sheet == null) {
			throw new IllegalStateException("Sheet is not intialized.");
		}
		return sheet;
	}

	public void setCommentHeight(Integer commentHeight) {
		if (commentHeight != null && commentHeight > 1) {
			this.commentHeight = commentHeight;
		}
	}

	public void setCommentWidth(Integer commentWidth) {
		if (commentWidth != null && commentWidth > 1) {
			this.commentWidth = commentWidth;
		}
	}

	public void setCommentAuthor(String commentAuthor) {
		if (commentAuthor != null && commentAuthor.trim().isEmpty() == false) {
			this.commentAuthor = commentAuthor;
		}
	}

	public boolean isSetupCellStylesForAllColumns() {
		return setupCellStylesForAllColumns;
	}

	public void setSetupCellStylesForAllColumns(boolean setupCellStylesForAllColumns) {
		this.setupCellStylesForAllColumns = setupCellStylesForAllColumns;
	}

	private boolean checkIfIsAppendedDataValidationNeccessary(DataValidation originalDv, int lastRowIndex) {
		CellRangeAddressList originalAl = originalDv.getRegions();
		int originalLastDataRow = 0;
		for (int i = 0; i < originalAl.countRanges(); i++) {
			CellRangeAddress cra = originalAl.getCellRangeAddress(i);
			if (cra.getLastRow() > originalLastDataRow) {
				originalLastDataRow = cra.getLastRow();
			}
		}
		return (originalLastDataRow < lastRowIndex);
	}

	private void createNewAppendingDataValidationAsCopy(Sheet sheet, DataValidation originalDv, int lastRowIndex) {
		CellRangeAddressList originalAl = originalDv.getRegions();
		CellRangeAddressList appendingAddressList = createNewAppendingCellRangeAddressList(originalAl, lastRowIndex);
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		DataValidation newValidation = dvHelper.createValidation(originalDv.getValidationConstraint(), appendingAddressList);
		newValidation.setSuppressDropDownArrow(originalDv.getSuppressDropDownArrow());
		newValidation.setShowErrorBox(originalDv.getShowErrorBox());
		newValidation.setShowPromptBox(originalDv.getShowPromptBox());
		newValidation.setEmptyCellAllowed(originalDv.getEmptyCellAllowed());
		newValidation.setErrorStyle(originalDv.getErrorStyle());
		if (originalDv.getValidationConstraint().getValidationType() == ValidationType.LIST) {
			// fix the problem for list constraints no formula will be applied in dvHelper.createValidation above
			DataValidationConstraint originalConstraint = originalDv.getValidationConstraint();
			DataValidationConstraint newConstraint = newValidation.getValidationConstraint();
			if (originalConstraint != null && newConstraint != null) {
				if (originalConstraint.getFormula1() != null) {
					newConstraint.setFormula1(originalConstraint.getFormula1());
				}
				if (originalConstraint.getFormula2() != null) {
					newConstraint.setFormula1(originalConstraint.getFormula2());
				}
				newConstraint.setOperator(originalConstraint.getOperator());
			}
		}
		String promptBoxText = originalDv.getPromptBoxText();
		String promptBoxTitle = originalDv.getPromptBoxTitle();
		String errorBoxText = originalDv.getErrorBoxText();
		String errorBoxTitle = originalDv.getErrorBoxTitle();
		if (promptBoxTitle != null && promptBoxText != null) {
			newValidation.createPromptBox(promptBoxTitle, promptBoxText);
		}
		if (errorBoxTitle != null && errorBoxText != null) {
			newValidation.createErrorBox(errorBoxTitle, errorBoxText);
		}
		sheet.addValidationData(newValidation);
	}
	
	private CellRangeAddressList createNewAppendingCellRangeAddressList(CellRangeAddressList originalAddressRangeList, int newLastRowIndex) {
		CellRangeAddressList extendedCellRangeAddressList = new CellRangeAddressList();
		for (CellRangeAddress ca : originalAddressRangeList.getCellRangeAddresses()) {
			extendedCellRangeAddressList.addCellRangeAddress(createAppendingCellRangeAddress(ca, newLastRowIndex));
		}
		return extendedCellRangeAddressList;
	}
	
	private CellRangeAddress createAppendingCellRangeAddress(CellRangeAddress originalAdressRange, int newLastRowIndex) {
		return new CellRangeAddress(originalAdressRange.getLastRow() + 1, newLastRowIndex, originalAdressRange.getFirstColumn(), originalAdressRange.getLastColumn());
	}
	
	private String getDataValidationConstraintDebugString(DataValidationConstraint dvc) {
		String constraint = "No constraint";
		if (dvc != null) {
			if (dvc.getValidationType() == ValidationType.LIST) {
				constraint = "List values: " + printArray(dvc.getExplicitListValues());
			} else if (dvc.getValidationType() == ValidationType.FORMULA) {
				constraint = "Formulas: #1: " + dvc.getFormula1() + " | #2: " + dvc.getFormula2();
			} else { 
				constraint = "Type: " + dvc.getValidationType();
			}
		}
		return constraint;
	}

	public void createDataValidationsForAppendedRows() {
		List<? extends DataValidation> dvs = sheet.getDataValidations();
		if (dvs != null && currentRow != null) {
			// check if there are data validations and we have at least one row written
			if (debug) {
				debug("Original list of DataValidations:");
				int i = 0;
				for (DataValidation dv : dvs) {
					debug("#" + i + " Adress range: " + dv.getRegions().getCellRangeAddresses()[0].formatAsString());
					debug("#" + i + "   Constraint: " + getDataValidationConstraintDebugString(dv.getValidationConstraint()));
					i++;
				}
			}
			info("Create new extended DataValidations (last written row: " + (currentRow.getRowNum() + 1) + "), number of validations: " + dvs.size());
			for (DataValidation dv : dvs) {
				if (checkIfIsAppendedDataValidationNeccessary(dv, currentRow.getRowNum())) {
					createNewAppendingDataValidationAsCopy(sheet, dv, currentRow.getRowNum());
				}
			}
			if (debug) {
				debug("New appended list of DataValidations:");
				dvs = sheet.getDataValidations();
				int i = 0;
				for (DataValidation dv : dvs) {
					debug("#" + i + " Adress range: " + dv.getRegions().getCellRangeAddresses()[0].formatAsString());
					debug("#" + i + "   Constraint: " + getDataValidationConstraintDebugString(dv.getValidationConstraint()));
					i++;
				}
			}
		}
	}

	public boolean isWriteZeroDateAsNull() {
		return writeZeroDateAsNull;
	}

	public void setWriteZeroDateAsNull(boolean writeZeroDateAsNull) {
		this.writeZeroDateAsNull = writeZeroDateAsNull;
	}

	public boolean isForbiddenWritingInProtectedCells() {
		return forbidWritingInProtectedCells;
	}

	public void setForbidWritingInProtectedCells(boolean forbidWritingInProtectedCells) {
		this.forbidWritingInProtectedCells = forbidWritingInProtectedCells;
	}

	public int getTemplateRowIndexForStyles() {
		return templateRowIndexForStyles;
	}

	public void setTemplateRowIndexForStyles(Integer templateRowIndexForStyles) {
		if (templateRowIndexForStyles != null) {
			this.templateRowIndexForStyles = templateRowIndexForStyles;
		} else {
			this.templateRowIndexForStyles = -1;
		}
	}
	
}
