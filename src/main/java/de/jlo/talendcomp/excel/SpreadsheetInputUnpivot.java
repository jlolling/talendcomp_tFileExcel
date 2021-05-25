package de.jlo.talendcomp.excel;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;

public class SpreadsheetInputUnpivot {
	
	/**
	 * This class holds the normalized values 
	 * @author jan.lolling@gmail.com
	 *
	 */
	public static class NormalizedRow {
		
		private int rowIndex = 0;
		private int originalColumnIndex = 0;
		private String headerName = null;
		private Object value = null;
		
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
		
		public String getHeaderName() {
			return headerName;
		}
		
		public void setHeaderName(String headerName) {
			if (headerName == null || headerName.trim().isEmpty()) {
				throw new IllegalArgumentException("headerName cannot be null or empty");
			}
			this.headerName = headerName;
		}
		
		public Object getValue() {
			return value;
		}
		
		public void setValue(Object value) {
			this.value = value;
		}
		
		@Override
		public boolean equals(Object o) {
			if (o instanceof NormalizedRow) {
				return ((NormalizedRow) o).headerName.equals(headerName);
			} else {
				return false;
			}
		}
		
		@Override
		public int hashCode() {
			return headerName.hashCode();
		}
		
	}

	private SpreadsheetInput spreadsheetInput = null;
	private int unpivotColumnRangeStartIndex = 0;
	private int unpivotColumnRangeEndIndex = 0;
	private int headerRowIndex = 0;
	private List<String> headerNames = new ArrayList<>();
	
	public int getUnpivotColumnRangeStartIndex() {
		return unpivotColumnRangeStartIndex;
	}

	public void setUnpivotColumnRangeStartIndex(int unpivotColumnRangeStartIndex) {
		this.unpivotColumnRangeStartIndex = unpivotColumnRangeStartIndex;
	}

	public int getUnpivotColumnRangeEndIndex() {
		return unpivotColumnRangeEndIndex;
	}

	public void setUnpivotColumnRangeEndIndex(int unpivotColumnRangeEndIndex) {
		this.unpivotColumnRangeEndIndex = unpivotColumnRangeEndIndex;
	}


	public void setSpreadsheetInput(SpreadsheetInput spreadsheetInput) {
		if (spreadsheetInput == null) {
			throw new IllegalArgumentException("The reference to the component tFileExcelSheetInput cannot be null");
		}
		this.spreadsheetInput = spreadsheetInput;
	}
	
	private void setupHeaderNames() {
		Row headerRow = spreadsheetInput.getRow(headerRowIndex);
		int cellIndex = unpivotColumnRangeStartIndex;
		boolean foundHeaderName = true;
		while (foundHeaderName) {
			try {
				String name = spreadsheetInput.getStringCellValue(cellIndex, true, true, false);
			} catch (Exception e) {
				
			}
		}
	}

	public int getHeaderRowIndex() {
		return headerRowIndex;
	}

	public void setHeaderRowIndex(Integer headerRowIndex) {
		if (headerRowIndex != null && headerRowIndex > 0) {
			this.headerRowIndex = headerRowIndex;
		}
	}
	
	
}