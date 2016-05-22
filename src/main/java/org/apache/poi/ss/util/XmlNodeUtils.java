package org.apache.poi.ss.util;

import java.util.Iterator;
import java.util.List;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.formula.EvaluationWorkbook.ExternalSheet;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.NamePtg;
import org.apache.poi.ss.formula.ptg.NameXPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * @author Guillaume de GENTILE (gentile_g at yahoo dot com)
 * jlolling: Fix added to follow the changed API
 */
public class XmlNodeUtils {
	
	private final XSSFWorkbook _wb;
	private final XSSFEvaluationWorkbook _fpwb;

	public XmlNodeUtils(XSSFWorkbook wb) {
		_wb = wb;
		_fpwb = XSSFEvaluationWorkbook.create(_wb);
	}

	/**
	 * update the sheet name in all references in this sheet
	 * @param sheetIndex current sheet index
	 * @param sheetname  the new sheet name (which was previously changed to this name)
	 */
	public void updateRelationsSheetName(final int sheetIndex, final String sheetname) {
		/**
		 * An instance of FormulaRenderingWorkbook that returns the new sheet name
		 */
		FormulaRenderingWorkbook frwb = new FormulaRenderingWorkbook() {

			@Override
			public ExternalSheet getExternalSheet(int externSheetIndex) {
				return _fpwb.getExternalSheet(externSheetIndex);
			}

			@Override
			public String resolveNameXText(NameXPtg nameXPtg) {
				return _fpwb.resolveNameXText(nameXPtg);
			}

			@Override
			public String getNameText(NamePtg namePtg) {
				return _fpwb.getNameText(namePtg);
			}

			@Override
			public String getSheetFirstNameByExternSheet(int externSheetIndex) {
				if (externSheetIndex == sheetIndex) {
					return sheetname;
				} else { 
					return _fpwb.getSheetFirstNameByExternSheet(externSheetIndex);
				}
			}

			@Override
			public String getSheetLastNameByExternSheet(int externSheetIndex) {
				if (externSheetIndex == sheetIndex) {
					return sheetname;
				} else { 
					return _fpwb.getSheetLastNameByExternSheet(externSheetIndex);
				}
			}
			
		};

		// update charts
		List<POIXMLDocumentPart> rels = _wb.getSheetAt(sheetIndex).getRelations();
		String oldSheetName = _wb.getSheetName(sheetIndex);

		// if the sheet being cloned has a drawing then update it
		XSSFDrawing dg = null;
		for (POIXMLDocumentPart r : rels) {
			// do not copy the drawing relationship, it will be re-created
			if (r instanceof XSSFDrawing) {
				dg = (XSSFDrawing) r;

				Iterator<XSSFChart> it = dg.getCharts().iterator();
				while (it.hasNext()) {
					XSSFChart chart = it.next();
					// System.out.println("chart = " + chart);
					CTChart c = chart.getCTChart();

					Node node1 = c.getDomNode();
					updateDomDocSheetReference(node1, frwb, oldSheetName);

					Node node2 = chart.getCTChartSpace().getDomNode();
					updateDomDocSheetReference(node2, frwb, oldSheetName);

				}
				continue;
			}
		}
	}

	/**
	 * Update sheet name in all formulas and named ranges.
	 * <p/>
	 * <p>
	 * The idea is to parse every formula and render it back to string with the
	 * updated sheet name.
	 * </p>
	 *
	 * @param rootNode
	 *            root node of the XML document
	 * @param sourceSheetIndex
	 *            the source sheet index
	 * @param targetSheetIndex
	 *            the target sheet index
	 */
	public void updateDomDocSheetReference(Node rootNode, final int sourceSheetIndex, final int targetSheetIndex) {
		final String sheetname = _wb.getSheetName(targetSheetIndex);
		/**
		 * An instance of FormulaRenderingWorkbook that returns
		 */
		FormulaRenderingWorkbook frwb = new FormulaRenderingWorkbook() {

			@Override
			public ExternalSheet getExternalSheet(int externSheetIndex) {
				return _fpwb.getExternalSheet(externSheetIndex);
			}

			@Override
			public String resolveNameXText(NameXPtg nameXPtg) {
				return _fpwb.resolveNameXText(nameXPtg);
			}

			@Override
			public String getNameText(NamePtg namePtg) {
				return _fpwb.getNameText(namePtg);
			}

			@Override
			public String getSheetFirstNameByExternSheet(int externSheetIndex) {
				if (externSheetIndex == sourceSheetIndex) {
					return sheetname;
				} else { 
					return _fpwb.getSheetFirstNameByExternSheet(externSheetIndex);
				}
			}

			@Override
			public String getSheetLastNameByExternSheet(int externSheetIndex) {
				if (externSheetIndex == sourceSheetIndex) {
					return sheetname;
				} else { 
					return _fpwb.getSheetLastNameByExternSheet(externSheetIndex);
				}
			}

		};

		String oldName = _wb.getSheetName(sourceSheetIndex);
		updateDomDocSheetReference(rootNode, frwb, oldName);
	}

	private void updateDomDocSheetReference(Node rootNode, FormulaRenderingWorkbook frwb, String oldName) {
		String value = rootNode.getNodeValue();
		// System.out.println(" " + rootNode.getNodeName() + " -> " +
		// rootNode.getNodeValue());
		if (value != null) {
			if (value.contains(oldName)) {
				XSSFName name1 = _wb.createName();
				name1.setRefersToFormula(value);
				updateName(name1, frwb);
				rootNode.setNodeValue(name1.getRefersToFormula());
				_wb.removeName(name1.getNameName());
			}
		}
		NodeList nl = rootNode.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++) {
			updateDomDocSheetReference(nl.item(i), frwb, oldName);
		}
	}

	/**
	 * Parse formula in the named range and re-assemble it back using the
	 * specified FormulaRenderingWorkbook.
	 *
	 * @param name
	 *            the name to update
	 * @param frwb
	 *            the formula rendering workbook that returns new sheet name
	 */
	private void updateName(XSSFName name, FormulaRenderingWorkbook frwb) {
		String formula = name.getRefersToFormula();
		if (formula != null) {
			int sheetIndex = name.getSheetIndex();
			Ptg[] ptgs = FormulaParser.parse(formula, _fpwb, FormulaType.NAMEDRANGE, sheetIndex);
			String updatedFormula = FormulaRenderer.toFormulaString(frwb, ptgs);
			if (formula.equals(updatedFormula) == false) {
				name.setRefersToFormula(updatedFormula);
			}
		}
	}

	public void printNode(Node rootNode, String spacer) {
		System.out.println(spacer + rootNode.getNodeName() + " -> " + rootNode.getNodeValue());
		NodeList nl = rootNode.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++) {
			printNode(nl.item(i), spacer + "   ");
		}
	}
		
}