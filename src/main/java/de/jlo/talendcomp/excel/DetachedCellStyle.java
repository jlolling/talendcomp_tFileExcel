package de.jlo.talendcomp.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public abstract class DetachedCellStyle implements CellStyle {
	
	private short dataFormat;
	private String dataFormatString;
	private Font font;
	private short fontIndex;
	private boolean hidden;
	private short indention;
	private short leftBorderColor;
	private short rightBorderColor;
	private short topBorderColor;
	private short buttomBorderColor;
	private boolean locked;
	private boolean quotePrefixed;
	private short rotation;
	private boolean shrinkToFit;
	private VerticalAlignment verticalAlignment;
	private boolean wrapText;
	private HorizontalAlignment horizontalAlignment;
	private short fillBackgroundColor;
	private short fillForegroundColor;
	private FillPatternType fillPatternType;
	
	@Override
	public void setDataFormat(short fmt) {
		this.dataFormat = fmt;
	}

	@Override
	public short getDataFormat() {
		return dataFormat;
	}

	@Override
	public String getDataFormatString() {
		return dataFormatString;
	}

	@Override
	public void setFont(Font font) {
		this.font = font;
	}
	
	public void setFontIndex(short index) {
		this.fontIndex = index;
	}

	@Override
	public short getFontIndex() {
		return fontIndex;
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	@Override
	public boolean getHidden() {
		return hidden;
	}

	@Override
	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	@Override
	public boolean getLocked() {
		return locked;
	}

	@Override
	public void setAlignment(HorizontalAlignment align) {
		this.horizontalAlignment = align;
	}

	@Override
	public HorizontalAlignment getAlignmentEnum() {
		return horizontalAlignment;
	}

	@Override
	public void setWrapText(boolean wrapped) {
		this.wrapText = wrapped;
	}

	@Override
	public boolean getWrapText() {
		return wrapText;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment align) {
		this.verticalAlignment = align;
	}

	@Override
	public VerticalAlignment getVerticalAlignmentEnum() {
		return verticalAlignment;
	}

	@Override
	public void setRotation(short rotation) {
		this.rotation = rotation;
	}

	@Override
	public short getRotation() {
		return rotation;
	}

	@Override
	public void setIndention(short indent) {
		this.indention = indent;
	}

	@Override
	public short getIndention() {
		return indention;
	}

	@Override
	public void setBorderLeft(BorderStyle border) {
		
	}

	@Override
	public BorderStyle getBorderLeftEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setBorderRight(BorderStyle border) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public BorderStyle getBorderRightEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setBorderTop(BorderStyle border) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public BorderStyle getBorderTopEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setBorderBottom(BorderStyle border) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public BorderStyle getBorderBottomEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setLeftBorderColor(short color) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getLeftBorderColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setRightBorderColor(short color) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getRightBorderColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setTopBorderColor(short color) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getTopBorderColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setBottomBorderColor(short color) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getBottomBorderColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void setFillPattern(FillPatternType fp) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFillPattern() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public FillPatternType getFillPatternEnum() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFillBackgroundColor(short bg) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFillBackgroundColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public Color getFillBackgroundColorColor() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setFillForegroundColor(short bg) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public short getFillForegroundColor() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public Color getFillForegroundColorColor() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void cloneStyleFrom(CellStyle source) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setShrinkToFit(boolean shrinkToFit) {
		this.shrinkToFit = shrinkToFit;
	}

	@Override
	public boolean getShrinkToFit() {
		return shrinkToFit;
	}

}
