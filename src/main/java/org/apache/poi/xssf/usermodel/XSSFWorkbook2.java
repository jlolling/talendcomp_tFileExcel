package org.apache.poi.xssf.usermodel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.util.XmlNodeUtils;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.w3c.dom.Node;

public class XSSFWorkbook2 extends XSSFWorkbook {

	private static POILogger logger = POILogFactory.getLogger(XSSFWorkbook2.class);
	
    public XSSFWorkbook2(InputStream is) throws IOException {
    	super(is);
    }

    public XSSFWorkbook2() throws IOException {
    	super();
    }

    /**
     * Validate sheet index
     *
     * @param index the index to validate
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
    */
    private void validateSheetIndex(int index) {
        int lastSheetIx = getNumberOfSheets() - 1;
        if (index < 0 || index > lastSheetIx) {
            String range = "(0.." +    lastSheetIx + ")";
            if (lastSheetIx == -1) {
                range = "(no sheets)";
            }
            throw new IllegalArgumentException("Sheet index ("
                    + index +") is out of range " + range);
        }
    }

    /**
     * Generate a valid sheet name based on the existing one. Used when cloning sheets.
     *
     * @param srcName the original sheet name to
     * @return clone sheet name
     */
    private String getUniqueSheetName(String srcName) {
        int uniqueIndex = 2;
        String baseName = srcName;
        int bracketPos = srcName.lastIndexOf('(');
        if (bracketPos > 0 && srcName.endsWith(")")) {
            String suffix = srcName.substring(bracketPos + 1, srcName.length() - ")".length());
            try {
                uniqueIndex = Integer.parseInt(suffix.trim());
                uniqueIndex++;
                baseName = srcName.substring(0, bracketPos).trim();
            } catch (NumberFormatException e) {
                // contents of brackets not numeric
            }
        }
        while (true) {
            // Try and find the next sheet name that is unique
            String index = Integer.toString(uniqueIndex++);
            String name;
            if (baseName.length() + index.length() + 2 < 31) {
                name = baseName + " (" + index + ")";
            } else {
                name = baseName.substring(0, 31 - index.length() - 2) + "(" + index + ")";
            }

            //If the sheet name is unique, then set it otherwise move on to the next number.
            if (getSheetIndex(name) == -1) {
                return name;
            }
        }
    }

    /**
	 * Create an XSSFSheet from an existing sheet in the XSSFWorkbook. The
	 * cloned sheet is a deep copy of the original.
	 *
	 * @return XSSFSheet representing the cloned sheet.
	 * @throws IllegalArgumentException
	 *             if the sheet index in invalid
	 * @throws POIXMLException
	 *             if there were errors when cloning
	 */
	@SuppressWarnings("deprecation")
	public XSSFSheet cloneSheet(int sheetNum) {
		validateSheetIndex(sheetNum);

		XSSFSheet srcSheet = getSheetAt(sheetNum);
		String srcName = srcSheet.getSheetName();
		String clonedName = getUniqueSheetName(srcName);

		XSSFSheet clonedSheet = createSheet(clonedName);
		try {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			srcSheet.write(out);
			clonedSheet.read(new ByteArrayInputStream(out.toByteArray()));
		} catch (IOException e) {
			throw new POIXMLException("Failed to clone sheet", e);
		}
		CTWorksheet ct = clonedSheet.getCTWorksheet();
		if (ct.isSetLegacyDrawing()) {
			logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
			ct.unsetLegacyDrawing();
		}
		if (ct.isSetPageSetup()) {
			logger.log(POILogger.WARN, "Cloning sheets with page setup is not yet supported.");
			ct.unsetPageSetup();
		}

		clonedSheet.setSelected(false);

		// copy sheet's relations
		List<POIXMLDocumentPart> rels = srcSheet.getRelations();
		// if the sheet being cloned has a drawing then remember it and
		// re-create tpoo
		XSSFDrawing dg = null;
		for (POIXMLDocumentPart r : rels) {
			// do not copy the drawing relationship, it will be re-created
			if (r instanceof XSSFDrawing) {
				dg = (XSSFDrawing) r;
				continue;
			}

			PackageRelationship rel = r.getPackageRelationship();
			clonedSheet.getPackagePart().addRelationship(rel.getTargetURI(), rel.getTargetMode(),
					rel.getRelationshipType());
			clonedSheet.addRelation(rel.getId(), r);
		}

		// clone the sheet drawing along with its relationships
		if (dg != null) {
			if (ct.isSetDrawing()) {
				// unset the existing reference to the drawing,
				// so that subsequent call of
				// clonedSheet.createDrawingPatriarch() will create a new one
				ct.unsetDrawing();
			}
			XSSFDrawing clonedDg = clonedSheet.createDrawingPatriarch();
			// copy drawing contents
			clonedDg.getCTDrawing().set(dg.getCTDrawing());

			// Clone drawing relations
			List<POIXMLDocumentPart> srcRels = srcSheet.createDrawingPatriarch().getRelations();
			for (POIXMLDocumentPart rel : srcRels) {
				if (rel instanceof XSSFChart) {
					XSSFChart chart = (XSSFChart) rel;
					try {
						// create new chart
						int chartNumber = getPackagePart().getPackage()
								.getPartsByContentType(XSSFRelation.CHART.getContentType()).size() + 1;
						XSSFChart c = (XSSFChart) clonedDg.createRelationship(XSSFRelation.CHART,
								XSSFFactory.getInstance(), chartNumber);

						// Instantiate new XmlNodeUtils
						XmlNodeUtils nodeUtils = new XmlNodeUtils(this);
						int clonedSheetNum = this.getSheetIndex(clonedSheet);

						// duplicate source CTChart
						// the new CTChart is still referencing the source
						// sheet!
						CTChart ctc = (CTChart) chart.getCTChart().copy();
						Node node = ctc.getPlotArea().getDomNode();
						nodeUtils.updateDomDocSheetReference(node, sheetNum, clonedSheetNum);
						c.getCTChart().set(ctc);

						// duplicate source CTChartSpace
						// the new CTChartSpace is still referencing the source
						// sheet!
						CTChartSpace ctcs = (CTChartSpace) chart.getCTChartSpace().copy();
						node = ctcs.getDomNode();
						nodeUtils.updateDomDocSheetReference(node, sheetNum, clonedSheetNum);
						c.getCTChartSpace().set(ctcs);

						// create new relation for the new chart
						PackageRelationship relation = c.getPackageRelationship();
						clonedDg.getPackagePart().addRelationship(relation.getTargetURI(), relation.getTargetMode(),
								relation.getRelationshipType(), relation.getId());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				} else {
					PackageRelationship relation = rel.getPackageRelationship();
					clonedSheet.createDrawingPatriarch().getPackagePart().addRelationship(relation.getTargetURI(),
							relation.getTargetMode(), relation.getRelationshipType(), relation.getId());
				}

			}
		}
		return clonedSheet;
	}

}
