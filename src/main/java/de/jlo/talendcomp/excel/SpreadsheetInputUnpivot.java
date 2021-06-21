package de.jlo.talendcomp.excel;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

public class SpreadsheetInputUnpivot {
	
	/**
	 * This class holds the normalized values 
	 * @author jan.lolling@gmail.com
	 *
	 */
	public static class NormalizedRow {
		
		private int rowIndex = 0;
		private int originalColumnIndex = 0;
		private Cell header = null;
		private Cell value = null;
		
		public int getRowIndex() {
			return rowIndex;
		}
		
		public void setRowIndex(int rowIndex) {
			this.rowIndex = rowIndex;
		}
		
		public int getOriginalColumnIndex() {
			return originalColumnIndex;
		}
		
		public void setOriginalColumnIndex(int originalColumnIndex) {
			this.originalColumnIndex = originalColumnIndex;
		}
		
		public Cell getHeader() {
			return header;
		}
		
		public void setHeader(Cell header) {
			if (header == null) {
				throw new IllegalArgumentException("headerName cannot be null");
			}
			this.header = header;
		}
		
		public Object getValue() {
			return value;
		}
		
		public void setValue(Cell value) {
			this.value = value;
		}
		
		@Override
		public boolean equals(Object o) {
			if (o instanceof NormalizedRow) {
				return ((NormalizedRow) o).header.equals(header);
			} else {
				return false;
			}
		}
		
		@Override
		public int hashCode() {
			return header.hashCode();
		}
		
	}

	private SpreadsheetInput spreadsheetInput = null;
	private int unpivotColumnIndexRangeStart = -1;
	private int unpivotColumnIndexRangeEnd = -1;
	private int headerRowIndex = 0;
	private List<Cell> headers = new ArrayList<>();
	private List<NormalizedRow> normalizedRows = new ArrayList<>();
	private int currentNormalizedRowIndex = 0;
	private NormalizedRow currentNormalizedRow = null;
	
	public int getUnpivotColumnRangeStartIndex() {
		return unpivotColumnIndexRangeStart;
	}

	/**
	 * Set the column index (0-based) at which the columns should be normalized
	 * @param unpivotColumnIndexRangeStart 0-based index of the start column
	 */
	public void setUnpivotColumnRangeStartIndex(Integer unpivotColumnIndexRangeStart) {
		if (unpivotColumnIndexRangeStart != null) {
			this.unpivotColumnIndexRangeStart = unpivotColumnIndexRangeStart;
		}
	}

	/**
	 * Set the column index (starts with "A") at which the columns should be normalized
	 * @param unpivotColumnIndexRangeStart 0-based index of the start column
	 */
	public void setUnpivotColumnRangeStartIndex(String unpivotColumnNameRangeStart) {
		if (unpivotColumnNameRangeStart != null && unpivotColumnNameRangeStart.trim().isEmpty() == false) {
			this.unpivotColumnIndexRangeStart = CellReference.convertColStringToIndex(unpivotColumnNameRangeStart);
		}
	}

	public int getUnpivotColumnRangeEndIndex() {
		return unpivotColumnIndexRangeEnd;
	}

	/**
	 * Set the column index (0-based) until the columns should be normalized
	 * @param unpivotColumnIndexRangeStart 0-based index of the start column
	 */
	public void setUnpivotColumnRangeEndIndex(Integer unpivotColumnIndexRangeEnd) {
		if (unpivotColumnIndexRangeEnd != null) {
			this.unpivotColumnIndexRangeEnd = unpivotColumnIndexRangeEnd + 1;
		}
	}

	/**
	 * Set the column index (starts with "A") until the columns should be normalized
	 * @param unpivotColumnIndexRangeStart 0-based index of the start column
	 */
	public void setUnpivotColumnRangeEndIndex(String unpivotColumnNameRangeEnd) {
		if (unpivotColumnNameRangeEnd != null) {
			this.unpivotColumnIndexRangeEnd = CellReference.convertColStringToIndex(unpivotColumnNameRangeEnd) + 1;
		}
	}

	/**
	 * initialize with the given spreadsheet input
	 * must be called within the main flow
	 * @param spreadsheetInput
	 * @throws Exception
	 */
	public void checkAndInitialize(SpreadsheetInput spreadsheetInput) throws Exception {
		if (isInitialized()) {
			return;
		}
		if (spreadsheetInput == null) {
			throw new IllegalArgumentException("The reference to the component tFileExcelSheetInput cannot be null");
		}
		this.spreadsheetInput = spreadsheetInput;
		setupHeader();
	}
	
	public boolean isInitialized() {
		return this.spreadsheetInput != null;
	}
	
	private void setupHeader() throws Exception {
		Row headerRow = spreadsheetInput.getRow(headerRowIndex);
		int cellIndex = 0;
		if (unpivotColumnIndexRangeStart > 0) {
			cellIndex = unpivotColumnIndexRangeStart;
		}
		boolean foundHeaderValue = false;
		while (true) {
			if (unpivotColumnIndexRangeEnd > 0 && cellIndex >= unpivotColumnIndexRangeEnd) {
				break;
			}
			Cell headerCell = headerRow.getCell(cellIndex);
			if (spreadsheetInput.isCellValueEmpty(headerCell)) {
				if (foundHeaderValue) {
					// we had found filled header and now we found an empty and stop here
					break;
				}
			} else {
				if (foundHeaderValue == false && unpivotColumnIndexRangeStart < 1) {
					// if we do not have found a header previously
					// we take the first filled header as start
					unpivotColumnIndexRangeStart = cellIndex;
				}
				foundHeaderValue = true;
				headers.add(headerCell);
			}
			cellIndex++;
		}
	}
	
	public int getNormalizedColumnCount() {
		return headers.size();
	}
	
	public int getNormalizeStartColumnIndex() {
		return unpivotColumnIndexRangeStart;
	}

	public void normalizeValuesOfCurrentRow() {
		normalizedRows.clear();
		currentNormalizedRowIndex = 0;
		for (int i = 0, n = headers.size(); i < n; i++) {
			Cell header = headers.get(i);
			Row currentDataRow = spreadsheetInput.getCurrentRow();
			int originalColumnIndex = unpivotColumnIndexRangeStart + i;
			Cell value = currentDataRow.getCell(originalColumnIndex);
			NormalizedRow nr = new NormalizedRow();
			nr.header = header;
			nr.value = value;
			nr.rowIndex = currentDataRow.getRowNum();
			nr.originalColumnIndex = originalColumnIndex;
			normalizedRows.add(nr);
		}
	}
	
	public int getCurrentOriginalRowIndex() {
		return currentNormalizedRow.rowIndex + 1;
	}
	
	public int getCurrentOriginalColumnIndex() {
		return currentNormalizedRow.originalColumnIndex;
	}

	public boolean nextNormalizedRow() {
		if (normalizedRows.isEmpty()) {
			currentNormalizedRow = null;
			return false;
		} else if (currentNormalizedRowIndex < normalizedRows.size()) {
			currentNormalizedRow = normalizedRows.get(currentNormalizedRowIndex++);
			return true;
		} else {
			currentNormalizedRow = null;
			return false;
		}
	}
	
	public String getCurrentHeaderAsString() throws Exception {
		try {
			return spreadsheetInput.getStringCellValue(currentNormalizedRow.header, -1);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized header as String. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Date getCurrentHeaderAsDate(String pattern) throws Exception {
		try {
			return spreadsheetInput.getDateCellValue(currentNormalizedRow.header, pattern);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized header as Date. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Double getCurrentHeaderAsDouble() throws Exception {
		try {
			return spreadsheetInput.getDoubleCellValue(currentNormalizedRow.header);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized header as Number. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Long getCurrentHeaderAsLong() throws Exception {
		Double v = getCurrentHeaderAsDouble();
		if (v != null) {
			return v.longValue();
		} else {
			return null;
		}
	}

	public Float getCurrentHeaderAsFloat() throws Exception {
		Double v = getCurrentHeaderAsDouble();
		if (v != null) {
			return v.floatValue();
		} else {
			return null;
		}
	}

	public Integer getCurrentHeaderAsInteger() throws Exception {
		Double v = getCurrentHeaderAsDouble();
		if (v != null) {
			return v.intValue();
		} else {
			return null;
		}
	}

	public Short getCurrentHeaderAsShort() throws Exception {
		Double v = getCurrentHeaderAsDouble();
		if (v != null) {
			return v.shortValue();
		} else {
			return null;
		}
	}

	public Boolean getCurrentHeaderAsBoolean() throws Exception {
		try {
			return spreadsheetInput.getBooleanCellValue(currentNormalizedRow.header);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized header as Boolean. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public BigDecimal getCurrentHeaderAsBigDecimal() throws Exception {
		Double v = getCurrentHeaderAsDouble();
		if (v != null) {
			return new BigDecimal(v);
		} else {
			return null;
		}
	}

	public String getCurrentValueAsString() throws Exception {
		try {
			return spreadsheetInput.getStringCellValue(currentNormalizedRow.value, -1);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized value as String. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Date getCurrentValueAsDate(String pattern) throws Exception {
		try {
			return spreadsheetInput.getDateCellValue(currentNormalizedRow.value, pattern);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized value as Date. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Double getCurrentValueAsDouble() throws Exception {
		try {
			return spreadsheetInput.getDoubleCellValue(currentNormalizedRow.value);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized value as Number. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	public Long getCurrentValueAsLong() throws Exception {
		Double v = getCurrentValueAsDouble();
		if (v != null) {
			return v.longValue();
		} else {
			return null;
		}
	}

	public Integer getCurrentValueAsInteger() throws Exception {
		Double v = getCurrentValueAsDouble();
		if (v != null) {
			return v.intValue();
		} else {
			return null;
		}
	}

	public Short getCurrentValueAsShort() throws Exception {
		Double v = getCurrentValueAsDouble();
		if (v != null) {
			return v.shortValue();
		} else {
			return null;
		}
	}

	public BigDecimal getCurrentValueAsBigDecimal() throws Exception {
		Double v = getCurrentValueAsDouble();
		if (v != null) {
			return new BigDecimal(v);
		} else {
			return null;
		}
	}

	public Float getCurrentValueAsFloat() throws Exception {
		Double v = getCurrentValueAsDouble();
		if (v != null) {
			return v.floatValue();
		} else {
			return null;
		}
	}

	public Boolean getCurrentValueAsBoolean() throws Exception {
		try {
			return spreadsheetInput.getBooleanCellValue(currentNormalizedRow.value);
		} catch (Exception e) {
			throw new Exception("Failed to get normalized value value as Boolean. row: " + (currentNormalizedRow.rowIndex + 1) + " column: " + currentNormalizedRow.originalColumnIndex + " Error: " + e.getMessage(), e);
		}
	}

	/**
	 * row index of the header row
	 * 0-based
	 * @return headerRowIndex
	 */
	public int getHeaderRowIndex() {
		return headerRowIndex;
	}

	/**
	 * row index of the header row
	 * 0-based
	 * @param headerRowIndex (0-based, for the GUI please take keep in mind the GUI sends 1-based)
	 */
	public void setHeaderRowIndex(Integer headerRowIndex) {
		if (headerRowIndex != null && headerRowIndex > 0) {
			this.headerRowIndex = headerRowIndex;
		}
	}
		
}