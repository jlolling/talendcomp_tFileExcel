package de.cimt.talendcomp.tfileexcelpoi;

import org.apache.poi.ss.usermodel.Sheet;

public class SpreadsheetList extends SpreadsheetFile {
	
	public int countSheets() throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		} else {
			return workbook.getNumberOfSheets();
		}
	}
	
	public String getSheetName(int sheetIndex) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		} else {
			return workbook.getSheetName(sheetIndex);
		}
	}
	
	public int getCountSheetRows(int sheetIndex) throws Exception {
		if (workbook == null) {
			throw new Exception("Workbook is not initialized!");
		} else {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			if (sheet != null) {
				return sheet.getLastRowNum();
			} else {
				return 0;
			}
		}
	}

}
