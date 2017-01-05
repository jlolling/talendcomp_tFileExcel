package de.cimt.talendcomp.tfileexcelpoi;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class StyleUtil {
	
	private Workbook workbook = null;
	
	public StyleUtil(Workbook workbook) {
		if (workbook == null) {
			throw new IllegalArgumentException("workbook cannot be null!");
		}
		this.workbook = workbook;
	}
	
	public String buildCSS(CellStyle cellStyle) {
		StringBuilder css = new StringBuilder();
		css.append(getAlignmentCSS(cellStyle));
		css.append(getIndentionCSS(cellStyle));
		css.append(getBorderCSS(cellStyle));
		css.append(getFillColorCSS(cellStyle));
		css.append(getFontCSS(cellStyle));
		return css.toString();
	}
	
	public String getFontCSS(CellStyle cellStyle) {
		StringBuilder css = new StringBuilder();
		short fontIndex = cellStyle.getFontIndex();
		Font font = workbook.getFontAt(fontIndex);
		if (font != null) {
			css.append("font-family:");
			css.append(font.getFontName());
			css.append(";");
			css.append("font-size:");
			css.append(font.getFontHeightInPoints());
			css.append("px;");
			if (font instanceof XSSFFont) {
				XSSFFont xf = (XSSFFont) font;
				XSSFColor color = xf.getXSSFColor();
				if (color != null) {
					css.append("color:");
					css.append(getColorCSSValue(color));
					css.append(";");
				}
			} else if (font instanceof HSSFFont) {
				HSSFFont hf = (HSSFFont) font;
				HSSFColor color = hf.getHSSFColor((HSSFWorkbook) workbook);
				if (color != null) {
					css.append("color:");
					css.append(getColorCSSValue(color));
					css.append(";");
				}
			}
			if (font.getBold()) {
				css.append("font-weight:bold;");
			}
			if (font.getItalic()) {
				css.append("font-style:italic;");
			}
			if (font.getStrikeout()) {
				css.append("text-decoration:line-through;");
			}
		}
		return css.toString();
	}
	
	public String getFillColorCSS(CellStyle cellStyle) {
		StringBuilder css = new StringBuilder();
		if (cellStyle instanceof XSSFCellStyle) {
			XSSFCellStyle style = (XSSFCellStyle) cellStyle;
			XSSFColor color = style.getFillForegroundXSSFColor();
			if (color != null) {
				css.append("background-color:");
				css.append(getColorCSSValue(color));
			}
		} else if (cellStyle instanceof HSSFCellStyle) {
			HSSFCellStyle style = (HSSFCellStyle) cellStyle;
			HSSFColor color = style.getFillForegroundColorColor();
			if (color != null) {
				css.append("background-color:");
				css.append(getColorCSSValue(color));
			}
		}
		if (css.length() > 0) {
			css.append(";");
		}
		return css.toString();
	}

	private String getColorCSSValue(HSSFColor color) {
		StringBuilder css = new StringBuilder();
		short[] ca = color.getTriplet();
		css.append("#");
		if (ca != null) {
			// we ignore here the opaque level in index 0
			for (int i = 0; i < ca.length; i++) {
				String hex = Integer.toHexString(ca[i] & 0xFF);
				if (hex.length() == 1) {
					hex = "0" + hex;
				}
				css.append(hex);
			}
		}
		return css.toString();
	}
	
	private String getColorCSSValue(XSSFColor color) {
		StringBuilder css = new StringBuilder();
		if (color.isRGB()) {
			byte[] ca = color.getRGB();
			css.append("#");
			if (ca != null) {
				// we ignore here the opaque level in index 0
				for (int i = 0; i < ca.length; i++) {
					String hex = Integer.toHexString(ca[i] & 0xFF);
					if (hex.length() == 1) {
						hex = "0" + hex;
					}
					css.append(hex);
				}
			}
		} else if (color.hasAlpha()) {
			byte[] ca = color.getARGB();
			css.append("#");
			if (ca != null) {
				// we ignore here the opaque level in index 0
				for (int i = 1; i < ca.length; i++) {
					String hex = Integer.toHexString(ca[i] & 0xFF);
					if (hex.length() == 1) {
						hex = "0" + hex;
					}
					css.append(hex);
				}
			}
		}
		return css.toString();
	}

	public String getAlignmentCSS(CellStyle cellStyle) {
		StringBuilder css = new StringBuilder();
		short align = cellStyle.getAlignment();
		switch (align) {
			case CellStyle.ALIGN_CENTER: css.append("text-align:center"); break;
			case CellStyle.ALIGN_FILL: css.append("text-align:fill"); break;
			case CellStyle.ALIGN_JUSTIFY: css.append("text-align:justified"); break;
			case CellStyle.ALIGN_LEFT: css.append("text-align:left"); break;
			case CellStyle.ALIGN_RIGHT: css.append("text-align:right"); break;
		}
		if (css.length() > 0) {
			css.append(";");
		}
		return css.toString();
	}
	
	public String getIndentionCSS(CellStyle cellStyle) {
		StringBuilder css = new StringBuilder();
		short indent = cellStyle.getIndention();
		if (indent > 0) {
			css.append("padding-left:");
			css.append(indent);
			css.append("px");
		}
		if (css.length() > 0) {
			css.append(";");
		}
		return css.toString();
	}

	public String getBorderCSS(CellStyle cellStyle) {
		short borderStyle = cellStyle.getBorderBottom();
		StringBuilder css = new StringBuilder();
		String side = "bottom";
		css.append(getBorderStyleCSS(borderStyle, side));
		side = "top";
		borderStyle = cellStyle.getBorderTop();
		css.append(getBorderStyleCSS(borderStyle, side));
		side = "left";
		borderStyle = cellStyle.getBorderLeft();
		css.append(getBorderStyleCSS(borderStyle, side));
		side = "right";
		borderStyle = cellStyle.getBorderRight();
		css.append(getBorderStyleCSS(borderStyle, side));
		return css.toString();
	}

	private String getBorderStyleCSS(short borderStyle, String side) {
		StringBuilder css = new StringBuilder();
		switch (borderStyle) {
			case CellStyle.BORDER_NONE: css.append("border-" + side + "-style:none"); break;
			case CellStyle.BORDER_DASH_DOT: css.append("border-" + side + "-style:dashed"); break;
			case CellStyle.BORDER_DASH_DOT_DOT: css.append("border-" + side + "-style:dashed"); break;
			case CellStyle.BORDER_DASHED: css.append("border-" + side + "-style:dashed"); break;
			case CellStyle.BORDER_DOTTED: css.append("border-" + side + "-style:dotted"); break;
			case CellStyle.BORDER_DOUBLE: css.append("border-" + side + "-style:double"); break;
			case CellStyle.BORDER_HAIR: css.append("border-" + side + "-style:solid;border-" + side + "-with:thin"); break;
			case CellStyle.BORDER_MEDIUM: css.append("border-" + side + "-style:solid;border-" + side + "-with:2px"); break;
			case CellStyle.BORDER_MEDIUM_DASH_DOT: css.append("border-" + side + "-style:dashed;border-" + side + "-with:1px"); break;
			case CellStyle.BORDER_MEDIUM_DASHED: css.append("border-" + side + "-style:dashed;border-" + side + "-with:1px"); break;
			case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT: css.append("border-" + side + "-style:dotted; border-" + side + "-with:1px"); break;
			case CellStyle.BORDER_THIN: css.append("border-" + side + "-style:solid;border-" + side + "-with:thin"); break;
			case CellStyle.BORDER_THICK: css.append("border-" + side + "-style:solid;border-" + side + "-with:2px"); break;
		}
		if (css.length() > 0) {
			css.append(";");
		}
		return css.toString();
	}
	
}
