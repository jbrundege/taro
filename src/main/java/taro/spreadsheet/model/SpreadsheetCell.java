package taro.spreadsheet.model;

import static java.lang.String.format;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFFont;

import taro.spreadsheet.TaroSpreadsheetException;


public class SpreadsheetCell {

	private SpreadsheetTab tab;
	private Cell cell;
	private SpreadsheetCellStyle style;

	SpreadsheetCell(SpreadsheetTab tab, Cell cell) {
		this.tab = tab;
		this.cell = cell;
	}

	public SpreadsheetCell setStyle(SpreadsheetCellStyle style) {
		CellStyle cellStyle = tab.registerStyle(style);
		cell.setCellStyle(cellStyle);
		this.style = style;
		return this;
	}

	/**
	 * Applies the given style to the current style for this cell, meaning non-null fields on the given style will be
	 * set on this cell's style, but null fields will be ignored.
	 * For instance, you could define a style that represents an 'invalid' cell and make the background color red and
	 * give it a red border. Then you could take any other style or cell and apply the invalid style to it. It would
	 * change the color to red and add the red border, but leave all other stylings (such as alignment, font, etc.) alone.
	 */
	public SpreadsheetCell applyStyle(SpreadsheetCellStyle toApply) {
		if (style == null) {
			setStyle(toApply);
		} else {
			SpreadsheetCellStyle newStyle = style.apply(toApply);
			setStyle(newStyle);
		}
		return this;
	}

	public SpreadsheetCell setValue(Object value) {
		if (value == null) {
			cell.setCellValue((String)null);
		} else if (value instanceof String) {
			if (((String) value).startsWith("=")) {
				cell.setCellFormula(((String) value).substring(1));
			} else {
				cell.setCellValue((String)value);
			}
		} else if (value instanceof Number) {
			Double num = ((Number)value).doubleValue();
			if (num.isNaN() || num.isInfinite()) {
				cell.setCellValue("");
			} else {
				cell.setCellValue(num);
			}
		} else if (value instanceof Date) {
			cell.setCellValue((Date)value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar)value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean)value);
		} else if (value instanceof RichTextString) {
			cell.setCellValue((RichTextString)value);
		} else {
			throw new TaroSpreadsheetException(format("Cannot set a %s [%s] as the spreadsheet cell content.", value.getClass().getSimpleName(), value.toString()));
		}
		return this;
	}

	public Cell getCell() {
		return cell;
	}

	public SpreadsheetCellStyle getStyle() {
		return style;
	}

	public SpreadsheetTab getTab() {
		return tab;
	}

	public Cell getPoiCell() {
		return cell;
	}

	public int getFontSizeInPoints() {
		if (style != null) {
			taro.spreadsheet.model.SpreadsheetFont font = style.getFont();
			if (font != null) {
				Integer size = font.getFontSizeInPoints();
				if (size != null) {
					return size;
				}
			}
		}
		return XSSFFont.DEFAULT_FONT_SIZE;
	}
}
