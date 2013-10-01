package taro.spreadsheet.model;

import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_RIGHT;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM;
import static org.apache.poi.ss.usermodel.CellStyle.VERTICAL_CENTER;

import java.awt.Color;

import org.apache.commons.lang.StringUtils;

public class SpreadsheetCellStyle {

	// Colors from the excel 'styles' box on the toolbar
	public static final Color COLOR_BAD = new Color(255, 199, 206);
	public static final Color COLOR_GOOD = new Color(198, 239, 206);
	public static final Color COLOR_NEUTRAL = new Color(255, 235, 156);
	public static final Color COLOR_NOTE = new Color(255, 255, 204);

	public static final SpreadsheetCellStyle DEFAULT = new SpreadsheetCellStyle();
	public static final SpreadsheetCellStyle CENTER = DEFAULT.withAlign(ALIGN_CENTER).withNumDecimals(0);
	public static final SpreadsheetCellStyle RIGHT = CENTER.withAlign(ALIGN_RIGHT);
	public static final SpreadsheetCellStyle LEFT = CENTER.withAlign(ALIGN_LEFT);
	public static final SpreadsheetCellStyle CENTER_ONE_DECIMAL = CENTER.withNumDecimals(1);

	public static final SpreadsheetCellStyle SUBTITLE = CENTER.withBold(true).withSurroundBorder(BORDER_MEDIUM);
	public static final SpreadsheetCellStyle TITLE = SUBTITLE.withFontSizeInPoints(14);
	public static final SpreadsheetCellStyle HEADER = SUBTITLE.withVerticalAlign(VERTICAL_CENTER).withWrapText(true);


	private SpreadsheetFont font;
	private Short align;
	private Short verticalAlign;
	private Short topBorder;
	private Short rightBorder;
	private Short bottomBorder;
	private Short leftBorder;
	private Color topBorderColor;
	private Color leftBorderColor;
	private Color bottomBorderColor;
	private Color rightBorderColor;
	private String dataFormatString;
	private Color backgroundColor;
	private Boolean locked;
	private Boolean hidden;
	private Boolean wrapText;
	private Integer indention;
	private Integer rotation;


	/**
	 * Returns a new style that applies the given style to this one, ignoring all null fields.
	 * For instance, you could define a style that represents an 'invalid' cell and make the background color red and
	 * give it a red border. Then you could take any other style or cell and apply the invalid style to it. It would
	 * change the color to red and add the red border, but leave all other stylings (such as alignment, font, etc.) alone.
	 */
	public SpreadsheetCellStyle apply(SpreadsheetCellStyle other) {
		SpreadsheetCellStyle copy = this.copy();
		apply(other, copy);
		return copy;
	}

	private void apply(SpreadsheetCellStyle source, SpreadsheetCellStyle destination) {
		if (source.align != null) destination.align = source.align;
		if (source.verticalAlign != null) destination.verticalAlign = source.verticalAlign;
		if (source.topBorder != null) destination.topBorder = source.topBorder;
		if (source.rightBorder != null) destination.rightBorder = source.rightBorder;
		if (source.bottomBorder != null) destination.bottomBorder = source.bottomBorder;
		if (source.leftBorder != null) destination.leftBorder = source.leftBorder;
		if (source.topBorderColor != null) destination.topBorderColor = source.topBorderColor;
		if (source.rightBorderColor != null) destination.rightBorderColor = source.rightBorderColor;
		if (source.bottomBorderColor != null) destination.bottomBorderColor = source.bottomBorderColor;
		if (source.leftBorderColor != null) destination.leftBorderColor = source.leftBorderColor;
		if (source.dataFormatString != null) destination.dataFormatString = source.dataFormatString;
		if (source.backgroundColor != null) destination.backgroundColor = source.backgroundColor;
		if (source.locked != null) destination.locked = source.locked;
		if (source.hidden != null) destination.hidden = source.hidden;
		if (source.wrapText != null) destination.wrapText = source.wrapText;
		if (source.indention != null) destination.indention = source.indention;
		if (source.rotation != null) destination.rotation = source.rotation;

		if (destination.font == null) {
			destination.font = source.font;
		} else {
			destination.font = destination.font.apply(source.font);
		}
	}

	public SpreadsheetCellStyle copy() {
		SpreadsheetCellStyle copy = new SpreadsheetCellStyle();
		apply(this, copy);
		return copy;
	}

	public SpreadsheetFont getFont() {
		return font;
	}

	public SpreadsheetCellStyle withFont(SpreadsheetFont font) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = font;
		return copy;
	}

	public SpreadsheetFont getOrCreateFont() {
		// don't change this classes font reference!
		if (font != null) return font;
		return new SpreadsheetFont();
	}

	public Short getAlign() {
		return align;
	}

	public SpreadsheetCellStyle withAlign(Short align) {
		SpreadsheetCellStyle copy = this.copy();
		copy.align = align;
		return copy;
	}

	public Color getBackgroundColor() {
		return backgroundColor;
	}

	public SpreadsheetCellStyle withBackgroundColor(Color backgroundColor) {
		SpreadsheetCellStyle copy = this.copy();
		copy.backgroundColor = backgroundColor;
		return copy;
	}

	public Boolean getBold() {
		if (font == null) return null;
		return font.getBold();
	}

	public SpreadsheetCellStyle withBold(Boolean bold) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withBold(bold);
		return copy;
	}

	public Short getBottomBorder() {
		return bottomBorder;
	}

	public SpreadsheetCellStyle withBottomBorder(Short bottomBorder) {
		SpreadsheetCellStyle copy = this.copy();
		copy.bottomBorder = bottomBorder;
		return copy;
	}

	public Color getBottomBorderColor() {
		return bottomBorderColor;
	}

	public SpreadsheetCellStyle withBottomBorderColor(Color bottomBorderColor) {
		SpreadsheetCellStyle copy = this.copy();
		copy.bottomBorderColor = bottomBorderColor;
		return copy;
	}

	public String getDataFormatString() {
		return dataFormatString;
	}

	public SpreadsheetCellStyle withDataFormatString(String dataFormatString) {
		SpreadsheetCellStyle copy = this.copy();
		copy.dataFormatString = dataFormatString;
		return copy;
	}

	public String getFontName() {
		if (font == null) return null;
		return font.getFontName();
	}

	public SpreadsheetCellStyle withFontName(String fontName) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withFontName(fontName);
		return copy;
	}

	public Integer getFontOffset() {
		if (font == null) return null;
		return font.getFontOffset();
	}

	public SpreadsheetCellStyle withFontOffset(Integer fontOffset) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withFontOffset(fontOffset);
		return copy;
	}

	public Integer getFontSizeInPoints() {
		if (font == null) return null;
		return font.getFontSizeInPoints();
	}

	public SpreadsheetCellStyle withFontSizeInPoints(Integer fontSizeInPoints) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withFontSizeInPoints(fontSizeInPoints);
		return copy;
	}

	public Boolean isHidden() {
		return hidden;
	}

	public SpreadsheetCellStyle withHidden(Boolean hidden) {
		SpreadsheetCellStyle copy = this.copy();
		copy.hidden = hidden;
		return copy;
	}

	public Boolean getItalic() {
		if (font == null) return null;
		return font.getItalic();
	}

	public SpreadsheetCellStyle withItalic(Boolean italic) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withItalic(italic);
		return copy;
	}

	public Short getLeftBorder() {
		return leftBorder;
	}

	public SpreadsheetCellStyle withLeftBorder(Short leftBorder) {
		SpreadsheetCellStyle copy = this.copy();
		copy.leftBorder = leftBorder;
		return copy;
	}

	public Color getLeftBorderColor() {
		return leftBorderColor;
	}

	public SpreadsheetCellStyle withLeftBorderColor(Color leftBorderColor) {
		SpreadsheetCellStyle copy = this.copy();
		copy.leftBorderColor = leftBorderColor;
		return copy;
	}

	public SpreadsheetCellStyle withSurroundBorder(Short border) {
		return this.withTopBorder(border).withLeftBorder(border).withBottomBorder(border).withRightBorder(border);
	}

	public Boolean getLocked() {
		return locked;
	}

	public SpreadsheetCellStyle withLocked(Boolean locked) {
		SpreadsheetCellStyle copy = this.copy();
		copy.locked = locked;
		return copy;
	}

	public Short getRightBorder() {
		return rightBorder;
	}

	public SpreadsheetCellStyle withRightBorder(Short rightBorder) {
		SpreadsheetCellStyle copy = this.copy();
		copy.rightBorder = rightBorder;
		return copy;
	}

	public Color getRightBorderColor() {
		return rightBorderColor;
	}

	public SpreadsheetCellStyle withRightBorderColor(Color rightBorderColor) {
		SpreadsheetCellStyle copy = this.copy();
		copy.rightBorderColor = rightBorderColor;
		return copy;
	}

	public Boolean getStrikeout() {
		if (font == null) return null;
		return font.getStrikeout();
	}

	public SpreadsheetCellStyle withStrikeout(Boolean strikeout) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withStrikeout(strikeout);
		return copy;
	}

	public Short getTopBorder() {
		return topBorder;
	}

	public SpreadsheetCellStyle withTopBorder(Short topBorder) {
		SpreadsheetCellStyle copy = this.copy();
		copy.topBorder = topBorder;
		return copy;
	}

	public Color getTopBorderColor() {
		return topBorderColor;
	}

	public SpreadsheetCellStyle withTopBorderColor(Color topBorderColor) {
		SpreadsheetCellStyle copy = this.copy();
		copy.topBorderColor = topBorderColor;
		return copy;
	}

	public Boolean getUnderline() {
		if (font == null) return null;
		return font.getUnderline();
	}

	public SpreadsheetCellStyle withUnderline(boolean underline) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withUnderline(underline);
		return copy;
	}

	public Boolean getDoubleUnderline() {
		if (font == null) return null;
		return font.getDoubleUnderline();
	}

	public SpreadsheetCellStyle withDoubleUnderline(boolean doubleUnderline) {
		SpreadsheetCellStyle copy = this.copy();
		copy.font = getOrCreateFont().withDoubleUnderline(doubleUnderline);
		return copy;
	}

	public Short getVerticalAlign() {
		return verticalAlign;
	}

	public SpreadsheetCellStyle withVerticalAlign(Short verticalAlign) {
		SpreadsheetCellStyle copy = this.copy();
		copy.verticalAlign = verticalAlign;
		return copy;
	}

	public Boolean getWrapText() {
		return wrapText;
	}

	public SpreadsheetCellStyle withWrapText(Boolean wrapText) {
		SpreadsheetCellStyle copy = this.copy();
		copy.wrapText = wrapText;
		return copy;
	}

	public Integer getIndention() {
		return indention;
	}

	public SpreadsheetCellStyle withIndention(Integer indention) {
		SpreadsheetCellStyle copy = this.copy();
		copy.indention = indention;
		return copy;
	}

	public Integer getRotation() {
		return rotation;
	}

	public SpreadsheetCellStyle withRotation(Integer rotation) {
		SpreadsheetCellStyle copy = this.copy();
		copy.rotation = rotation;
		return copy;
	}

	public SpreadsheetCellStyle withNumDecimals(int numDecimals) {
		SpreadsheetCellStyle copy = this.copy();
		String dataFormat;
		if (numDecimals < 1) {
			dataFormat = "0";
		} else {
			dataFormat = StringUtils.rightPad("0.", numDecimals + 2, "0");
		}
		copy.dataFormatString = dataFormat;
		return copy;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (!(o instanceof SpreadsheetCellStyle)) return false;

		SpreadsheetCellStyle that = (SpreadsheetCellStyle) o;

		if (align != null ? !align.equals(that.align) : that.align != null) return false;
		if (backgroundColor != null ? !backgroundColor.equals(that.backgroundColor) : that.backgroundColor != null)
			return false;
		if (bottomBorder != null ? !bottomBorder.equals(that.bottomBorder) : that.bottomBorder != null) return false;
		if (bottomBorderColor != null ? !bottomBorderColor.equals(that.bottomBorderColor) : that.bottomBorderColor != null)
			return false;
		if (dataFormatString != null ? !dataFormatString.equals(that.dataFormatString) : that.dataFormatString != null)
			return false;
		if (font != null ? !font.equals(that.font) : that.font != null) return false;
		if (hidden != null ? !hidden.equals(that.hidden) : that.hidden != null) return false;
		if (indention != null ? !indention.equals(that.indention) : that.indention != null) return false;
		if (leftBorder != null ? !leftBorder.equals(that.leftBorder) : that.leftBorder != null) return false;
		if (leftBorderColor != null ? !leftBorderColor.equals(that.leftBorderColor) : that.leftBorderColor != null)
			return false;
		if (locked != null ? !locked.equals(that.locked) : that.locked != null) return false;
		if (rightBorder != null ? !rightBorder.equals(that.rightBorder) : that.rightBorder != null) return false;
		if (rightBorderColor != null ? !rightBorderColor.equals(that.rightBorderColor) : that.rightBorderColor != null)
			return false;
		if (rotation != null ? !rotation.equals(that.rotation) : that.rotation != null) return false;
		if (topBorder != null ? !topBorder.equals(that.topBorder) : that.topBorder != null) return false;
		if (topBorderColor != null ? !topBorderColor.equals(that.topBorderColor) : that.topBorderColor != null)
			return false;
		if (verticalAlign != null ? !verticalAlign.equals(that.verticalAlign) : that.verticalAlign != null)
			return false;
		if (wrapText != null ? !wrapText.equals(that.wrapText) : that.wrapText != null) return false;

		return true;
	}

	@Override
	public int hashCode() {
		int result = font != null ? font.hashCode() : 0;
		result = 31 * result + (align != null ? align.hashCode() : 0);
		result = 31 * result + (verticalAlign != null ? verticalAlign.hashCode() : 0);
		result = 31 * result + (topBorder != null ? topBorder.hashCode() : 0);
		result = 31 * result + (rightBorder != null ? rightBorder.hashCode() : 0);
		result = 31 * result + (bottomBorder != null ? bottomBorder.hashCode() : 0);
		result = 31 * result + (leftBorder != null ? leftBorder.hashCode() : 0);
		result = 31 * result + (topBorderColor != null ? topBorderColor.hashCode() : 0);
		result = 31 * result + (leftBorderColor != null ? leftBorderColor.hashCode() : 0);
		result = 31 * result + (bottomBorderColor != null ? bottomBorderColor.hashCode() : 0);
		result = 31 * result + (rightBorderColor != null ? rightBorderColor.hashCode() : 0);
		result = 31 * result + (dataFormatString != null ? dataFormatString.hashCode() : 0);
		result = 31 * result + (backgroundColor != null ? backgroundColor.hashCode() : 0);
		result = 31 * result + (locked != null ? locked.hashCode() : 0);
		result = 31 * result + (hidden != null ? hidden.hashCode() : 0);
		result = 31 * result + (wrapText != null ? wrapText.hashCode() : 0);
		result = 31 * result + (indention != null ? indention.hashCode() : 0);
		result = 31 * result + (rotation != null ? rotation.hashCode() : 0);
		return result;
	}

}
