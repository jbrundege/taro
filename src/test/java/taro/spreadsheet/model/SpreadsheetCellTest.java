package taro.spreadsheet.model;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.not;
import static org.hamcrest.CoreMatchers.notNullValue;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.closeTo;

import java.awt.Color;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.hamcrest.Matchers;
import org.junit.Before;
import org.junit.Test;

public class SpreadsheetCellTest {

	private SpreadsheetCell cell;

	@Before
	public void setup() {
		SpreadsheetWorkbook workbook = new SpreadsheetWorkbook();
		SpreadsheetTab tab = workbook.createTab("test");
		cell = tab.getCell("A1");
	}

	@Test
	public void setValueWithString_SetsAStringValueOnTheCell() {
		cell.setValue("A String");
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_STRING));
		assertThat(cell.getPoiCell().getStringCellValue(), is("A String"));
	}

	@Test
	public void setValueWithStringFormula_SetsAFormulaOnTheCell() {
		cell.setValue("=B1*C1");	// formula is any string starting with an equals (=) sign
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_FORMULA));
		assertThat(cell.getPoiCell().getCellFormula(), is("B1*C1"));
	}

	@Test
	public void setValueWithShort_SetsANumericValueOnTheCell() {
		cell.setValue((short)12);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), is(12.0));
	}

	@Test
	public void setValueWithInteger_SetsANumericValueOnTheCell() {
		cell.setValue(12);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), is(12.0));
	}

	@Test
	public void setValueWithLong_SetsANumericValueOnTheCell() {
		cell.setValue(12L);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), is(12.0));
	}

	@Test
	public void setValueWithFloat_SetsANumericValueOnTheCell() {
		cell.setValue(12f);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), is(12.0));
	}

	@Test
	public void setValueWithDouble_SetsANumericValueOnTheCell() {
		cell.setValue(12.0);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), is(12.0));
	}

	@Test
	public void setValueWithDate_SetsADateValueOnTheCell() {
		Date date = new Date();
		double excelDateNumber = DateUtil.getExcelDate(date);

		cell.setValue(date);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), closeTo(excelDateNumber, 0.001));
	}

	@Test
	public void setValueWithCalendar_SetsADateValueOnTheCell() {
		Calendar calendar = Calendar.getInstance();
		double excelDateNumber = DateUtil.getExcelDate(calendar.getTime());

		cell.setValue(calendar);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_NUMERIC));
		assertThat(cell.getPoiCell().getNumericCellValue(), closeTo(excelDateNumber, 0.001));
	}

	@Test
	public void setValueWithBoolean_SetsABooleanValueOnTheCell() {
		cell.setValue(true);
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_BOOLEAN));
		assertThat(cell.getPoiCell().getBooleanCellValue(), is(true));
	}

	@Test
	public void setValueWithNull_SetsABlankValueOnTheCell() {
		cell.setValue("A String to value to start with");
		cell.setValue(null);	// wipe out the string value
		assertThat(cell.getPoiCell().getCellType(), is(Cell.CELL_TYPE_BLANK));
		assertThat(cell.getPoiCell().getStringCellValue(), is(""));
	}

	@Test
	public void cellsStyle_IsNullUntilSet() {
		assertThat(cell.getStyle(), is(nullValue()));

		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withAlign(CellStyle.ALIGN_CENTER);
		cell.setStyle(cellStyle);

		assertThat(cell.getStyle(), is(notNullValue()));
		assertThat(cell.getStyle(), is(cellStyle));
	}

	@Test
	public void applyStyle_MergesAStyleOntoACell() {
		SpreadsheetCellStyle originalStyle = new SpreadsheetCellStyle().withAlign(CellStyle.ALIGN_CENTER).withBold(true)
				.withTopBorder(CellStyle.BORDER_MEDIUM)
				.withLeftBorder(CellStyle.BORDER_MEDIUM)
				.withBottomBorder(CellStyle.BORDER_MEDIUM)
				.withRightBorder(CellStyle.BORDER_MEDIUM);

		cell.setStyle(originalStyle);

		SpreadsheetCellStyle styleToApply = new SpreadsheetCellStyle()
				.withTopBorderColor(Color.RED)
				.withLeftBorderColor(Color.RED)
				.withBottomBorderColor(Color.RED)
				.withRightBorderColor(Color.RED);

		// Method Under Test
		cell.applyStyle(styleToApply);

		SpreadsheetCellStyle currentStyle = cell.getStyle();
		assertThat(currentStyle.getAlign(), is(CellStyle.ALIGN_CENTER));
		assertThat(currentStyle.getBold(), is(true));
		assertThat(currentStyle.getTopBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(currentStyle.getLeftBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(currentStyle.getBottomBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(currentStyle.getRightBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(currentStyle.getTopBorderColor(), is(Color.RED));
		assertThat(currentStyle.getLeftBorderColor(), is(Color.RED));
		assertThat(currentStyle.getBottomBorderColor(), is(Color.RED));
		assertThat(currentStyle.getRightBorderColor(), is(Color.RED));
	}

	@Test
	public void getFontSizeInPoints_ReturnsFontSizeIfAStyleWithAFontExists() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withFontSizeInPoints(15);
		cell.setStyle(cellStyle);
		assertThat(cell.getFontSizeInPoints(), is(15));
	}

	@Test
	public void getFontSizeInPoints_ReturnsTheDefaultFontSizeIfStyleOrFontIsNull() {
		assertThat(cell.getStyle(), is(nullValue()));
		assertThat(cell.getFontSizeInPoints(), is((int)XSSFFont.DEFAULT_FONT_SIZE));

		cell.setStyle(new SpreadsheetCellStyle());
		assertThat(cell.getStyle(), is(notNullValue()));
		assertThat(cell.getStyle().getFont(), is(nullValue()));
		assertThat(cell.getFontSizeInPoints(), is((int) XSSFFont.DEFAULT_FONT_SIZE));

		cell.setStyle(new SpreadsheetCellStyle().withFontSizeInPoints(17));
		assertThat(cell.getStyle(), is(notNullValue()));
		assertThat(cell.getStyle().getFont(), is(notNullValue()));
		assertThat(cell.getFontSizeInPoints(), is(not((int) XSSFFont.DEFAULT_FONT_SIZE)));
		assertThat(cell.getFontSizeInPoints(), is(17));
	}

}
