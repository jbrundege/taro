package taro.spreadsheet.model;

public class SpreadsheetFont {

	private String fontName;
	private Integer fontOffset;
	private Boolean bold;
	private Boolean italic;
	private Boolean underline;
	private Boolean doubleUnderline;
	private Boolean strikeout;
	private Integer sizeInPoints;

	public SpreadsheetFont copy() {
		SpreadsheetFont copy = new SpreadsheetFont();
		apply(this, copy);
		return copy;
	}
	
	public SpreadsheetFont apply(SpreadsheetFont other) {
		SpreadsheetFont copy = this.copy();
		if (other != null) {
			apply(other, copy);
		}
		return copy;
	}

	private void apply(SpreadsheetFont source, SpreadsheetFont destination) {
		if (source.fontName != null) destination.fontName = source.fontName;
		if (source.fontOffset != null) destination.fontOffset = source.fontOffset;
		if (source.bold != null) destination.bold = source.bold;
		if (source.italic != null) destination.italic = source.italic;
		if (source.underline != null) destination.underline = source.underline;
		if (source.doubleUnderline != null) destination.doubleUnderline = source.doubleUnderline;
		if (source.strikeout != null) destination.strikeout = source.strikeout;
		if (source.sizeInPoints != null) destination.sizeInPoints = source.sizeInPoints;
	}
	
	public String getFontName() {
		return fontName;
	}

	public SpreadsheetFont withFontName(String fontName) {
		SpreadsheetFont copy = this.copy();
		copy.fontName = fontName;
		return copy;
	}

	public Integer getFontOffset() {
		return fontOffset;
	}

	public SpreadsheetFont withFontOffset(Integer fontOffset) {
		SpreadsheetFont copy = this.copy();
		copy.fontOffset = fontOffset;
		return copy;
	}

	public Boolean getBold() {
		return bold;
	}

	public SpreadsheetFont withBold(Boolean bold) {
		SpreadsheetFont copy = this.copy();
		copy.bold = bold;
		return copy;
	}

	public Boolean getDoubleUnderline() {
		return doubleUnderline;
	}

	public SpreadsheetFont withDoubleUnderline(Boolean doubleUnderline) {
		SpreadsheetFont copy = this.copy();
		copy.doubleUnderline = doubleUnderline;
		copy.underline = null;
		return copy;
	}

	public Integer getFontSizeInPoints() {
		return sizeInPoints;
	}

	public SpreadsheetFont withFontSizeInPoints(Integer heightInPoints) {
		SpreadsheetFont copy = this.copy();
		copy.sizeInPoints = heightInPoints;
		return copy;
	}

	public Boolean getItalic() {
		return italic;
	}

	public SpreadsheetFont withItalic(Boolean italic) {
		SpreadsheetFont copy = this.copy();
		copy.italic = italic;
		return copy;
	}

	public Boolean getStrikeout() {
		return strikeout;
	}

	public SpreadsheetFont withStrikeout(Boolean strikeout) {
		SpreadsheetFont copy = this.copy();
		copy.strikeout = strikeout;
		return copy;
	}

	public Boolean getUnderline() {
		return underline;
	}

	public SpreadsheetFont withUnderline(boolean underline) {
		SpreadsheetFont copy = this.copy();
		copy.underline = underline;
		copy.doubleUnderline = null;
		return copy;
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (!(o instanceof SpreadsheetFont)) return false;

		SpreadsheetFont that = (SpreadsheetFont) o;

		if (bold != null ? !bold.equals(that.bold) : that.bold != null) return false;
		if (doubleUnderline != null ? !doubleUnderline.equals(that.doubleUnderline) : that.doubleUnderline != null)
			return false;
		if (fontName != null ? !fontName.equals(that.fontName) : that.fontName != null) return false;
		if (fontOffset != null ? !fontOffset.equals(that.fontOffset) : that.fontOffset != null) return false;
		if (italic != null ? !italic.equals(that.italic) : that.italic != null) return false;
		if (sizeInPoints != null ? !sizeInPoints.equals(that.sizeInPoints) : that.sizeInPoints != null) return false;
		if (strikeout != null ? !strikeout.equals(that.strikeout) : that.strikeout != null) return false;
		if (underline != null ? !underline.equals(that.underline) : that.underline != null) return false;

		return true;
	}

	@Override
	public int hashCode() {
		int result = fontName != null ? fontName.hashCode() : 0;
		result = 31 * result + (fontOffset != null ? fontOffset.hashCode() : 0);
		result = 31 * result + (bold != null ? bold.hashCode() : 0);
		result = 31 * result + (italic != null ? italic.hashCode() : 0);
		result = 31 * result + (underline != null ? underline.hashCode() : 0);
		result = 31 * result + (doubleUnderline != null ? doubleUnderline.hashCode() : 0);
		result = 31 * result + (strikeout != null ? strikeout.hashCode() : 0);
		result = 31 * result + (sizeInPoints != null ? sizeInPoints.hashCode() : 0);
		return result;
	}

}
