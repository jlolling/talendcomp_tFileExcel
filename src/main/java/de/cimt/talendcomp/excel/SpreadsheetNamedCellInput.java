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
package de.cimt.talendcomp.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellReference;


public class SpreadsheetNamedCellInput extends SpreadsheetFile {

	private int namedCellCount = 0;
	private int currentNamedCellIndex = 0;
	private Cell currentNamedCell = null;
	private String valueClass;
	private String cellName;

	public void retrieveNamedCellCount() {
		namedCellCount = workbook.getNumberOfNames();
	}
	
	public boolean readNextNamedCell() {
		if (workbook ==  null) {
			throw new IllegalStateException("workbook is not initialized");
		}
		if (namedCellCount == 0) {
			return false;
		} else {
			if (currentNamedCellIndex < namedCellCount) {
				Name name = workbook.getNameAt(currentNamedCellIndex);
				cellName = name.getNameName();
				currentNamedCell = getNamedCell(name);
				currentNamedCellIndex++;
				return true;
			} else {
				currentNamedCellIndex++;
				return false;
			}
		}
	}
	
	public int getNumberOfNamedCells() {
		return namedCellCount;
	}
	
	public int getCurrentCellIndex() {
		return currentNamedCellIndex - 1;
	}
	
	public Object getCellValue() {
		if (currentNamedCell != null) { // cell.getCellTypeEnum() == CellType.BLANK
			if (currentNamedCell.getCellTypeEnum() == CellType.BLANK) {
				valueClass = null;
				return null;
			} else if (currentNamedCell.getCellTypeEnum() == CellType.BOOLEAN) {
				valueClass = "java.lang.Boolean";
				return currentNamedCell.getBooleanCellValue();
			} else if (currentNamedCell.getCellTypeEnum() == CellType.ERROR) {
				valueClass = null;
				return null;
			} else if (currentNamedCell.getCellTypeEnum() == CellType.FORMULA) {
				valueClass = "java.lang.String";
				return getDataFormatter().formatCellValue(currentNamedCell, getFormulaEvaluator());
			} else if (currentNamedCell.getCellTypeEnum() == CellType.NUMERIC) {
				if (DateUtil.isCellDateFormatted(currentNamedCell)) {
					valueClass = "java.util.Date";
					return currentNamedCell.getDateCellValue();
				} else {
					valueClass = "java.lang.Double";
					return currentNamedCell.getNumericCellValue();
				}
			} else if (currentNamedCell.getCellTypeEnum() == CellType.STRING) {
				valueClass = "java.lang.String";
				return currentNamedCell.getStringCellValue();
			} else {
				valueClass = null;
				return null;
			}
		} else {
			valueClass = null;
			return null;
		}
	}
	
	public String getValueClass() {
		return valueClass;
	}
	
	public String getCellName() {
		return cellName;
	}
	
	public int getCellRowIndex() {
		if (currentNamedCell != null) {
			return currentNamedCell.getRowIndex() + 1;
		} else {
			return -1;
		}
	}
	
	public int getCellColumnIndex() {
		if (currentNamedCell != null) {
			return currentNamedCell.getColumnIndex();
		} else {
			return -1;
		}
	}

	public String getCellExcelReference() {
		if (currentNamedCell != null) {
	    	CellReference reference = new CellReference(currentNamedCell.getRowIndex(), currentNamedCell.getColumnIndex(), true, true);
	    	return reference.formatAsString();
		} else {
			return null;
		}
	}
	
	public String getCellSheetName() {
		if (currentNamedCell != null) {
			return currentNamedCell.getSheet().getSheetName();
		} else {
			return null;
		}
	}
	
}
