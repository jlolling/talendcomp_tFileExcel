/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package org.apache.poi.ss.examples;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.cimt.talendcomp.tfileexcelpoi.SpreadsheetOutput;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Excel Conditional Formatting -- Examples
 *
 * <p>
 *   Based on the code snippets from http://www.contextures.com/xlcondformat03.html
 * </p>
 *
 * @author Yegor Kozlov
 */
public class ConditionalFormats {

    public static void main(String[] args) throws IOException {
    	testStyleCopyAndCF();
    	/*
        Workbook wb;

        if(args.length > 0 && args[0].equals("-xls")) wb = new HSSFWorkbook();
        else wb = new XSSFWorkbook();

        sameCell(wb.createSheet("Same Cell"));
        multiCell(wb.createSheet("MultiCell"));
        errors(wb.createSheet("Errors"));
        hideDupplicates(wb.createSheet("Hide Dups"));
        formatDuplicates(wb.createSheet("Duplicates"));
        inList(wb.createSheet("In List"));
        expiry(wb.createSheet("Expiry"));
        shadeAlt(wb.createSheet("Shade Alt"));
        shadeBands(wb.createSheet("Shade Bands"));

        // Write the output to a file
        String file = "/var/testdata/excel/cf-poi.xls";
        if(wb instanceof XSSFWorkbook) file += "x";
        FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        out.close();
        */
    }

    /**
     * Highlight cells based on their values
     */
    static void sameCell(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue(84);
        sheet.createRow(1).createCell(0).setCellValue(74);
        sheet.createRow(2).createCell(0).setCellValue(50);
        sheet.createRow(3).createCell(0).setCellValue(51);
        sheet.createRow(4).createCell(0).setCellValue(49);
        sheet.createRow(5).createCell(0).setCellValue(41);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Cell Value Is   greater than  70   (Blue Fill)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "70");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Condition 2: Cell Value Is  less than      50   (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "50");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.GREEN.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1, rule2);

        sheet.getRow(0).createCell(2).setCellValue("<== Condition 1: Cell Value Is greater than 70 (Blue Fill)");
        sheet.getRow(4).createCell(2).setCellValue("<== Condition 2: Cell Value Is less than 50 (Green Fill)");
    }

    /**
     * Highlight multiple cells based on a formula
     */
    static void multiCell(Sheet sheet) {
        // header row
        Row row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Units");
        row0.createCell(1).setCellValue("Cost");
        row0.createCell(2).setCellValue("Total");

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(71);
        row1.createCell(1).setCellValue(29);
        row1.createCell(2).setCellValue(2059);

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue(85);
        row2.createCell(1).setCellValue(29);
        row2.createCell(2).setCellValue(2059);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue(71);
        row3.createCell(1).setCellValue(29);
        row3.createCell(2).setCellValue(2059);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =$B2>75   (Blue Fill)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("$A2>75");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:C4")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(4).setCellValue("<== Condition 1: Formula Is =$B2>75   (Blue Fill)");
    }

    /**
     *  Use Excel conditional formatting to check for errors,
     *  and change the font colour to match the cell colour.
     *  In this example, if formula result is  #DIV/0! then it will have white font colour.
     */
    static void errors(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue(84);
        sheet.createRow(1).createCell(0).setCellValue(0);
        sheet.createRow(2).createCell(0).setCellFormula("ROUND(A1/A2,0)");
        sheet.createRow(3).createCell(0).setCellValue(0);
        sheet.createRow(4).createCell(0).setCellFormula("ROUND(A6/A4,0)");
        sheet.createRow(5).createCell(0).setCellValue(41);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =ISERROR(C2)   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("ISERROR(A1)");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(1).setCellValue("<== The error in this cell is hidden. Condition: Formula Is   =ISERROR(C2)   (White Font)");
        sheet.getRow(4).createCell(1).setCellValue("<== The error in this cell is hidden. Condition: Formula Is   =ISERROR(C2)   (White Font)");
    }

    /**
     * Use Excel conditional formatting to hide the duplicate values,
     * and make the list easier to read. In this example, when the table is sorted by Region,
     * the second (and subsequent) occurrences of each region name will have white font color.
     */
    static void hideDupplicates(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("City");
        sheet.createRow(1).createCell(0).setCellValue("Boston");
        sheet.createRow(2).createCell(0).setCellValue("Boston");
        sheet.createRow(3).createCell(0).setCellValue("Chicago");
        sheet.createRow(4).createCell(0).setCellValue("Chicago");
        sheet.createRow(5).createCell(0).setCellValue("New York");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("A2=A1");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(1).createCell(1).setCellValue("<== the second (and subsequent) " +
                "occurences of each region name will have white font colour.  " +
                "Condition: Formula Is   =A2=A1   (White Font)");
    }

    /**
     * Use Excel conditional formatting to highlight duplicate entries in a column.
     */
    static void formatDuplicates(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Code");
        sheet.createRow(1).createCell(0).setCellValue(4);
        sheet.createRow(2).createCell(0).setCellValue(3);
        sheet.createRow(3).createCell(0).setCellValue(6);
        sheet.createRow(4).createCell(0).setCellValue(3);
        sheet.createRow(5).createCell(0).setCellValue(5);
        sheet.createRow(6).createCell(0).setCellValue(8);
        sheet.createRow(7).createCell(0).setCellValue(0);
        sheet.createRow(8).createCell(0).setCellValue(2);
        sheet.createRow(9).createCell(0).setCellValue(8);
        sheet.createRow(10).createCell(0).setCellValue(6);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A11")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(1).setCellValue("<== Duplicates numbers in the column are highlighted.  " +
                "Condition: Formula Is =COUNTIF($A$2:$A$11,A2)>1   (Blue Font)");
    }

    /**
     * Use Excel conditional formatting to highlight items that are in a list on the worksheet.
     */
    static void inList(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Codes");
        sheet.createRow(1).createCell(0).setCellValue("AA");
        sheet.createRow(2).createCell(0).setCellValue("BB");
        sheet.createRow(3).createCell(0).setCellValue("GG");
        sheet.createRow(4).createCell(0).setCellValue("AA");
        sheet.createRow(5).createCell(0).setCellValue("FF");
        sheet.createRow(6).createCell(0).setCellValue("XX");
        sheet.createRow(7).createCell(0).setCellValue("CC");

        sheet.getRow(0).createCell(2).setCellValue("Valid");
        sheet.getRow(1).createCell(2).setCellValue("AA");
        sheet.getRow(2).createCell(2).setCellValue("BB");
        sheet.getRow(3).createCell(2).setCellValue("CC");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("COUNTIF($C$2:$C$4,A2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A8")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(3).setCellValue("<== Use Excel conditional formatting to highlight items that are in a list on the worksheet");
    }

    /**
     *  Use Excel conditional formatting to highlight payments that are due in the next thirty days.
     *  In this example, Due dates are entered in cells A2:A4.
     */
    static void expiry(Sheet sheet) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setDataFormat((short)BuiltinFormats.getBuiltinFormat("d-mmm"));

        sheet.createRow(0).createCell(0).setCellValue("Date");
        sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
        sheet.createRow(2).createCell(0).setCellFormula("A2+1");
        sheet.createRow(3).createCell(0).setCellFormula("A3+1");

        for(int rownum = 1; rownum <= 3; rownum++) sheet.getRow(rownum).getCell(0).setCellStyle(style);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A4")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(0).createCell(1).setCellValue("Dates within the next 30 days are highlighted");
    }

    /**
     * Use Excel conditional formatting to shade alternating rows on the worksheet
     */
    static void shadeAlt(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.createRow(0).createCell(1).setCellValue("Shade Alternating Rows");
        sheet.createRow(1).createCell(1).setCellValue("Condition: Formula Is  =MOD(ROW(),2)   (Light Green Fill)");
    }

    /**
     * You can use Excel conditional formatting to shade bands of rows on the worksheet. 
     * In this example, 3 rows are shaded light grey, and 3 are left with no shading.
     * In the MOD function, the total number of rows in the set of banded rows (6) is entered.
     */
    static void shadeBands(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),6)<3");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.createRow(0).createCell(1).setCellValue("Shade Bands of Rows");
        sheet.createRow(1).createCell(1).setCellValue("Condition: Formula Is  =MOD(ROW(),6)<2   (Light Grey Fill)");
    }
    
    static void fillTemplate() {
    	File ft = new File("/var/testdata/excel/cf-template.xlsx");
    	try {
			Workbook workbook = new XSSFWorkbook(new FileInputStream(ft));
			Sheet sheet = workbook.cloneSheet(0);
			CellStyle cs1 = null;
			CellStyle cs2 = null;
			for (int r = 1; r < 1000; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					row = sheet.createRow(r);
				}
				Cell cell = row.getCell(0);
				if (cell == null) {
					cell = row.createCell(0);
				}
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				if (r == 1) {
					cs1 = cell.getCellStyle();
				}
				if (r == 2) {
					cs2 = cell.getCellStyle();
				}
				if (r > 2 && r % 2 > 0) {
					cell.setCellStyle(cs1);
				} else if (r > 2 && r % 2 == 0) {
					cell.setCellStyle(cs2);
				}
				cell.setCellValue(r);
			}
			SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
			ConditionalFormatting cf = find(scf, 1, 0);
			if (cf != null) {
		        CellRangeAddress[] regions = {
		                CellRangeAddress.valueOf(
		                		new CellReference(1, 0, false, false).formatAsString() +
		                		":" +
		                		new CellReference(1000, 0, false, false).formatAsString()
		                		)
		                	
		        };
				int numRules = cf.getNumberOfRules();
				for (int i = 0; i < numRules; i++) {
					scf.addConditionalFormatting(regions, cf.getRule(i));
				}
			}
			workbook.write(new FileOutputStream("/var/testdata/cf-result.xlsx"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
    
    private static ConditionalFormatting find(SheetConditionalFormatting scf, int row, int col) {
    	ConditionalFormatting cf = null;
		int numCF = scf.getNumConditionalFormattings();
		for (int i = 0; i < numCF; i++) {
			cf = scf.getConditionalFormattingAt(i);
			CellRangeAddress[] crArray = cf.getFormattingRanges();
			for (CellRangeAddress cra : crArray) {
				if (cra.isFullRowRange() == false) {
					if (cra.isInRange(row, col)) {
						return cf;
					}
				} else {
					System.out.println("isFullRowRange");
				}
			}
		}
		return null;
    }
    
    private static void testStyleCopyAndCF() {
    	File ft = new File("/var/testdata/excel/cf-template.xlsx");
    	try {
			Workbook workbook = new XSSFWorkbook(new FileInputStream(ft));
	    	SpreadsheetOutput so = new SpreadsheetOutput();
	    	so.setOutputFile("/var/testdata/excel/cf-result.xlsx");
	    	so.setWorkbook(workbook);
	    	so.createCopy(0, "Jan");
	    	so.initializeSheet();
	    	so.setRowStartIndex(1);
	    	so.setColumnStart(0);
			so.setDataFormat(1, "#,##0.0");
	    	so.setReuseExistingStyles(true);
	    	so.setReuseExistingStylesAlternating(true);
	    	Object[] dataset = new Object[2];
	    	System.out.println("Fill data...");
	    	for (double r = 1; r < 1000; r++) {
	    		dataset[0] = new java.util.Date();
	    		dataset[1] = r + 0.2345678;
	    		so.writeRow(dataset);
	    	}
	    	System.out.println("Extend cell ranges...");
	    	so.extendCellRangesForConditionalFormattings();
	    	so.writeWorkbook();
    	} catch (Exception e) {
    		e.printStackTrace();
    	}
    }

}