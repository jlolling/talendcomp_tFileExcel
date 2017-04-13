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
