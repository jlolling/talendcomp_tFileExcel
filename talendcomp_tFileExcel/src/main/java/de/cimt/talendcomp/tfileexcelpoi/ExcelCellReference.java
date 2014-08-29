package de.cimt.talendcomp.tfileexcelpoi;

/**
 * @author  lolling.jan  help to handle with Excel-formated cell position
 */
public class ExcelCellReference {
    
    private short columnIndex = 0;
    private int rowIndex = 0;
    private boolean isAbsoluteColumnPos = false;
    private boolean isAbsoluteRowPos = false;
    final static int COL_RADIX = 26;
    
    /**
     * column index start with 0
     * @return  column index
     */
    public short getColumnIndex() {
        return columnIndex;
    }
    
    /**
     * set column index starting at 0
     * @param  columnIndex
     */
    public void setColumnIndex(short columnIndex) {
        this.columnIndex = columnIndex;
    }
    
    /**
     * return row index (starting with 0)
     * @return  row index
     */
    public int getRowIndex() {
        return rowIndex;
    }
    
    /**
     * set row index starting with 0
     * @param rowIndex  row index
     */
    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
    
    /**
     * type of column position
     * @return  true if absolute
     */
    public boolean isAbsoluteColumnPos() {
        return isAbsoluteColumnPos;
    }
    
    /**
     * type of row position
     * @return  true if absolute
     */
    public boolean isAbsoluteRowPos() {
        return isAbsoluteRowPos;
    }
    
    /**
     * type of column position
     * @param isAbsoluteColumnPos true if absolute
     */
    public void setIsAbsoluteColumnPos(boolean isAbsoluteColumnPos) {
        this.isAbsoluteColumnPos = isAbsoluteColumnPos;
    }
    
    /**
     * type of row position
     * @param isAbsoluteRowPos true if absolute
     */
    public void setIsAbsoluteRowPos(boolean isAbsoluteRowPos) {
        this.isAbsoluteRowPos = isAbsoluteRowPos;
    }
    
    public void incrementColumnIndex() {
        columnIndex++;
    }
    
    public void incrementRowIndex() {
        rowIndex++;
    }
    
    /**
     * translates an instance of ExcelCellPosition in a String like Excel-Format
     * @param pos position
     * @return String like Excel
     */
    public static String translateToExcelPos(ExcelCellReference pos) {
        return translateToExcelPos(pos.getColumnIndex(),
                                   pos.getRowIndex(),
                                   pos.isAbsoluteColumnPos,
                                   pos.isAbsoluteRowPos);
    }

    /**
     * translate to Excel-Position-String
     * @param columnIndex column-index (0-based)
     * @param rowIndex    row-index    (0-based)
     * @param absoluteColumn true if absolute column
     * @param absoluteRow    true if absolute row
     * @return cell position in Excel format e.g. "$AB$23" or "G7"
     */
    public static String translateToExcelPos(int columnIndex, int rowIndex, boolean absoluteColumn, boolean absoluteRow) {
        char[] buf = new char[16];
        int lastPos = 0;
        // columns
        if (absoluteColumn) {
            buf[lastPos++] = '$';
        }
        if (columnIndex >= COL_RADIX) {
            buf[lastPos++] = (char) ('A' + (short) ((columnIndex / COL_RADIX) - 1));
            buf[lastPos++] = (char) ('A' + (short) (columnIndex % COL_RADIX));
        } else {
            buf[lastPos++] = (char) ('A' + (short) columnIndex);
        }
        if (absoluteRow) {
            buf[lastPos++] = '$';
        }
        // rows
        rowIndex++;
        int startRow = lastPos;
        while (rowIndex >= 10) {
            buf[lastPos++] = (char) ('0' + (rowIndex % 10));
            rowIndex = rowIndex / 10;
        }
        buf[lastPos] = (char) ('0' + rowIndex);
        // change direction
        char c;
        for (int i = 0; i < ((lastPos - startRow) + 1) / 2; i++) {
            c = buf[lastPos - i];
            buf[lastPos - i] = buf[startRow + i];
            buf[startRow + i] = c;
        }
        return String.valueOf(buf, 0, lastPos + 1);
    }
    
    /**
     * wandelt einen Spaltenindex in eine Excel-Spaltenbezeichnung um
     * @param columnIndex Index der Spalte
     * @param absoluteColumn true wenn absolute Spalte
     * @return Spaltenbezeichnung
     */
    public String translateColumnIndexToExcel(int columnIndex, boolean absoluteColumn) {
        char[] buf = new char[3]; // max 65536 row in Excel allowed
        int lastPos = 0;
        // columns
        if (absoluteColumn) {
            buf[lastPos++] = '$';
        }
        if (columnIndex >= COL_RADIX) {
            buf[lastPos++] = (char) ('A' + (short) ((columnIndex / COL_RADIX) - 1));
            buf[lastPos++] = (char) ('A' + (short) (columnIndex % COL_RADIX));
        } else {
            buf[lastPos++] = (char) ('A' + (short) columnIndex);
        }
        return String.valueOf(buf, 0, lastPos + 1);
    }
    
    /**
     * erstellt den Index einer Spalte passend zum Namen
     * @param columnName Spaltenname
     * @return Spaltenindex
     */
    public static int parseCellColumnName(String columnName) {
        int columnIndex = 0;
        char c;
        for (int i = 0; i < columnName.length(); i++) {
            c = columnName.charAt(i);
            if (Character.isUpperCase(c)) {
                columnIndex = columnIndex * COL_RADIX + (c - 'A') + 1;
            } else if (Character.isDigit(c)) {
                throw new NumberFormatException("invalid char in "+columnName+" at pos="+i+":"+c);
            } else if (columnIndex == 0) {
                throw new NumberFormatException("unable to parse "+columnName+" at pos="+i+":"+c);
            }
        }
        return (columnIndex - 1);
    }
    
    public static ExcelCellReference parseCellPos(String cellPosStr) throws NumberFormatException {
        return parseCellPos(null, cellPosStr);
    }
    
    /**
     * converts an Excel-cell-position into an ExcelCellPosition instance
     * @param ref reference to change xy values
     * @param cellPosStr Excel-cell position like "$AB$23" or "G7"
     * @return position
     * @throws NumberFormatException
     */
    public static ExcelCellReference parseCellPos(ExcelCellReference pos, String cellPosStr) throws NumberFormatException {
        int rowIndex = 0;
        int columnIndex = 0;
        boolean inColumn = true;
        boolean isAbsoluteColumn = false;
        boolean isAbsoluteRow = false;
        char c;
        for (int i = 0; i < cellPosStr.length(); i++) {
            c = cellPosStr.charAt(i);
            if (inColumn) {
	            if (Character.isUpperCase(c)) {
	                columnIndex = columnIndex * COL_RADIX + (c - 'A') + 1;
	            } else if (c == '$') {
	                if (columnIndex == 0) {
	                    isAbsoluteColumn = true;
	                } else {
		                inColumn = false;
		                isAbsoluteRow = true;
	                }
	            } else if (Character.isDigit(c)) {
	                // start of row
	                rowIndex = (c - '1') + 1;
	                inColumn = false;
	            } else if (columnIndex == 0) {
	                throw new NumberFormatException("unable to parse "+cellPosStr+" at pos="+i+":"+c);
	            }
            } else {
                if (Character.isDigit(c)) {
                    rowIndex = rowIndex * 10 + (c - '1') + 1;
                } else {
                    throw new NumberFormatException("unable to parse "+cellPosStr+" at pos="+i+":"+c);
                }
            }
        }
        if (rowIndex == 0) {
            throw new NumberFormatException("unable to parse "+cellPosStr+" 0 as row index is not allowed");
        } else if (columnIndex == 0) {
            throw new NumberFormatException("unable to parse "+cellPosStr+" 0 as column index is not allowed");
        }
        if (pos == null) {
            pos = new ExcelCellReference();
        }
        pos.setColumnIndex((short) (columnIndex - 1));
        pos.setRowIndex(rowIndex - 1);
        pos.setIsAbsoluteRowPos(isAbsoluteRow);
        pos.setIsAbsoluteColumnPos(isAbsoluteColumn);
        return pos;
    }
    
    public String toString() {
        if (isAbsoluteColumnPos && isAbsoluteRowPos) {
            return "$"+columnIndex+":$"+rowIndex;
        } else if (isAbsoluteColumnPos) {
            return "$"+columnIndex+":"+rowIndex;
        } else if (isAbsoluteRowPos) {
            return columnIndex+":$"+rowIndex;
        } else {
            return columnIndex+":"+rowIndex;
        }
    }
    
}
