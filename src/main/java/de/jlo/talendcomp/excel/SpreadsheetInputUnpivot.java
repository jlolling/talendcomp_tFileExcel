package de.jlo.talendcomp.excel;

import org.apache.poi.ss.usermodel.Row;

public class SpreadsheetInputUnpivot {
	
	private SpreadsheetInput input = null;
	private int unpivotColumnRangeStartIndex = 0;
	private int unpivotColumnRangeEndIndex = 0;
	private Row headerRow = null;
	private Row currentRow = null;
	
	public void setInput(SpreadsheetInput input) {
		this.input = input;
	}

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

	public Row getHeaderRow() {
		return headerRow;
	}

	public void setHeaderRow(Row headerRow) {
		this.headerRow = headerRow;
	}

	public Row getCurrentRow() {
		return currentRow;
	}

	public void setCurrentRow(Row currentRow) {
		this.currentRow = currentRow;
	}

	
	
}
