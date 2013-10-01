package taro.spreadsheet.model;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.not;
import static org.hamcrest.CoreMatchers.notNullValue;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.CoreMatchers.sameInstance;
import static org.hamcrest.MatcherAssert.assertThat;

import java.awt.Color;

import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Test;

public class SpreadsheetCellStyleTest {


	@Test
	public void aNewSpreadsheetCellStyle_HasAllNullProperties() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle();
		assertThat(cellStyle.getBold(), nullValue());
		assertThat(cellStyle.getItalic(), nullValue());
		assertThat(cellStyle.getStrikeout(), nullValue());
		assertThat(cellStyle.getUnderline(), nullValue());
		assertThat(cellStyle.getDoubleUnderline(), nullValue());
		assertThat(cellStyle.getFontName(), nullValue());
		assertThat(cellStyle.getFontOffset(), nullValue());
		assertThat(cellStyle.getFontSizeInPoints(), nullValue());

		assertThat(cellStyle.getAlign(), nullValue());
		assertThat(cellStyle.getBackgroundColor(), nullValue());
		assertThat(cellStyle.getBottomBorder(), nullValue());
		assertThat(cellStyle.getBottomBorderColor(), nullValue());
		assertThat(cellStyle.getLeftBorder(), nullValue());
		assertThat(cellStyle.getLeftBorderColor(), nullValue());
		assertThat(cellStyle.getTopBorder(), nullValue());
		assertThat(cellStyle.getTopBorderColor(), nullValue());
		assertThat(cellStyle.getRightBorder(), nullValue());
		assertThat(cellStyle.getRightBorderColor(), nullValue());
		assertThat(cellStyle.getDataFormatString(), nullValue());
		assertThat(cellStyle.getIndention(), nullValue());
		assertThat(cellStyle.getLocked(), nullValue());
		assertThat(cellStyle.getRotation(), nullValue());
		assertThat(cellStyle.getVerticalAlign(), nullValue());
		assertThat(cellStyle.getWrapText(), nullValue());
		assertThat(cellStyle.getFont(), nullValue());
	}

	@Test
	public void SpreadsheetCellStyle_IsImmutable() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle();
		assertThat(cellStyle.withBold(true), is(not(cellStyle)));
		assertThat(cellStyle.withDoubleUnderline(true), is(not(cellStyle)));
		assertThat(cellStyle.withFontName("Courier"), is(not(cellStyle)));
		assertThat(cellStyle.withFontOffset(1), is(not(cellStyle)));
		assertThat(cellStyle.withFontSizeInPoints(14), is(not(cellStyle)));
		assertThat(cellStyle.withItalic(true), is(not(cellStyle)));
		assertThat(cellStyle.withStrikeout(true), is(not(cellStyle)));
		assertThat(cellStyle.withUnderline(true), is(not(cellStyle)));

		assertThat(cellStyle.withAlign(CellStyle.ALIGN_CENTER), is(not(cellStyle)));
		assertThat(cellStyle.withBackgroundColor(Color.BLUE), is(not(cellStyle)));
		assertThat(cellStyle.withBottomBorder(CellStyle.BORDER_MEDIUM), is(not(cellStyle)));
		assertThat(cellStyle.withBottomBorderColor(Color.BLUE), is(not(cellStyle)));
		assertThat(cellStyle.withLeftBorder(CellStyle.BORDER_MEDIUM), is(not(cellStyle)));
		assertThat(cellStyle.withLeftBorderColor(Color.BLUE), is(not(cellStyle)));
		assertThat(cellStyle.withTopBorder(CellStyle.BORDER_MEDIUM), is(not(cellStyle)));
		assertThat(cellStyle.withTopBorderColor(Color.BLUE), is(not(cellStyle)));
		assertThat(cellStyle.withRightBorder(CellStyle.BORDER_MEDIUM), is(not(cellStyle)));
		assertThat(cellStyle.withRightBorderColor(Color.BLUE), is(not(cellStyle)));
		assertThat(cellStyle.withDataFormatString("0.00"), is(not(cellStyle)));
		assertThat(cellStyle.withIndention(1), is(not(cellStyle)));
		assertThat(cellStyle.withLocked(true), is(not(cellStyle)));
		assertThat(cellStyle.withRotation(1), is(not(cellStyle)));
		assertThat(cellStyle.withVerticalAlign(CellStyle.VERTICAL_TOP), is(not(cellStyle)));
		assertThat(cellStyle.withWrapText(true), is(not(cellStyle)));

		assertThat(cellStyle.copy(), not(sameInstance(cellStyle)));
		assertThat(cellStyle.apply(new SpreadsheetCellStyle()), not(sameInstance(cellStyle)));
	}

	@Test
	public void apply_TransfersAllPropertiesToCopy() {
		SpreadsheetCellStyle src = new SpreadsheetCellStyle()
				.withBold(true).withFontName("Courier")
				.withFontOffset(1)
				.withFontSizeInPoints(14)
				.withItalic(true)
				.withStrikeout(true)
				.withUnderline(true)
				.withAlign(CellStyle.ALIGN_CENTER)
				.withBackgroundColor(Color.BLUE)
				.withBottomBorder(CellStyle.BORDER_MEDIUM)
				.withBottomBorderColor(Color.BLUE)
				.withLeftBorder(CellStyle.BORDER_MEDIUM)
				.withLeftBorderColor(Color.BLUE)
				.withTopBorder(CellStyle.BORDER_MEDIUM)
				.withTopBorderColor(Color.BLUE)
				.withRightBorder(CellStyle.BORDER_MEDIUM)
				.withRightBorderColor(Color.BLUE)
				.withDataFormatString("0.00")
				.withIndention(1)
				.withLocked(true)
				.withRotation(1)
				.withVerticalAlign(CellStyle.VERTICAL_TOP)
				.withWrapText(true);

		SpreadsheetCellStyle dest = new SpreadsheetCellStyle();

		// Method Under Test
		SpreadsheetCellStyle applied = dest.apply(src);

		assertThat(applied.getBold(), is(true));
		assertThat(applied.getItalic(), is(true));
		assertThat(applied.getStrikeout(), is(true));
		assertThat(applied.getUnderline(), is(true));
		assertThat(applied.getFontName(), is("Courier"));
		assertThat(applied.getFontOffset(), is(1));
		assertThat(applied.getFontSizeInPoints(), is(14));

		assertThat(applied.getAlign(), is(CellStyle.ALIGN_CENTER));
		assertThat(applied.getBackgroundColor(), is(Color.BLUE));
		assertThat(applied.getBottomBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(applied.getBottomBorderColor(), is(Color.BLUE));
		assertThat(applied.getLeftBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(applied.getLeftBorderColor(), is(Color.BLUE));
		assertThat(applied.getTopBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(applied.getTopBorderColor(), is(Color.BLUE));
		assertThat(applied.getRightBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(applied.getRightBorderColor(), is(Color.BLUE));
		assertThat(applied.getDataFormatString(), is("0.00"));
		assertThat(applied.getIndention(), is(1));
		assertThat(applied.getLocked(), is(true));
		assertThat(applied.getRotation(), is(1));
		assertThat(applied.getVerticalAlign(), is(CellStyle.VERTICAL_TOP));
		assertThat(applied.getWrapText(), is(true));
	}

	@Test
	public void apply_OverwritesNonNullProperties() {
		SpreadsheetCellStyle src = new SpreadsheetCellStyle().withBold(true).withLeftBorder(CellStyle.BORDER_THIN);
		SpreadsheetCellStyle dest = new SpreadsheetCellStyle().withIndention(1).withLeftBorder(CellStyle.BORDER_MEDIUM);

		assertThat(src.getBold(), is(true));
		assertThat(src.getLeftBorder(), is(CellStyle.BORDER_THIN));
		assertThat(src.getIndention(), nullValue());

		assertThat(dest.getBold(), nullValue());
		assertThat(dest.getLeftBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(dest.getIndention(), is(1));

		// Method Under Test
		SpreadsheetCellStyle applied = dest.apply(src);

		// src was set, so overwrite the null value on dest
		assertThat(applied.getBold(), is(true));

		// src was set, so overwrite the previously set value on dest
		assertThat(applied.getLeftBorder(), is(CellStyle.BORDER_THIN));

		// src was not set, so do not overwrite the existing set value on dest
		assertThat(applied.getIndention(), is(1));
	}

	@Test
	public void copy_ReturnsNewInstanceWithSameProperties() {
		SpreadsheetCellStyle original = new SpreadsheetCellStyle().withBold(true).withDataFormatString("0.00")
				.withRightBorder(CellStyle.BORDER_THIN);

		assertThat(original.getBold(), is(true));
		assertThat(original.getDataFormatString(), is("0.00"));
		assertThat(original.getRightBorder(), is(CellStyle.BORDER_THIN));
		assertThat(original.getUnderline(), nullValue());
		assertThat(original.getLeftBorder(), nullValue());

		// Method Under Test
		SpreadsheetCellStyle copy = original.copy();

		assertThat(copy, not(sameInstance(original)));

		assertThat(copy.getBold(), is(true));
		assertThat(copy.getDataFormatString(), is("0.00"));
		assertThat(copy.getRightBorder(), is(CellStyle.BORDER_THIN));
		assertThat(copy.getUnderline(), nullValue());
		assertThat(copy.getLeftBorder(), nullValue());
	}

	@Test
	public void equals_IsTrueWhenDifferentFontsHaveTheSameProperties() {
		SpreadsheetCellStyle one = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);
		SpreadsheetCellStyle two = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);

		assertThat(one, not(sameInstance(two)));
		assertThat(one.equals(two), is(true));
	}

	@Test
	public void equals_IsFalseIfAnyPropertyIsDifferent() {
		SpreadsheetCellStyle one = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);
		SpreadsheetCellStyle two = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("#,##0").withAlign(CellStyle.ALIGN_CENTER);
		// one and two differ only in data format string

		assertThat(one, not(sameInstance(two)));
		assertThat(one.equals(two), is(false));
	}

	@Test
	public void hashCode_IsSameWhenDifferentFontsHaveTheSameProperties() {
		SpreadsheetCellStyle one = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);
		SpreadsheetCellStyle two = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);

		assertThat(one, not(sameInstance(two)));
		assertThat(one.hashCode(), is(two.hashCode()));
	}

	@Test
	public void hashCode_IsDifferentIfAnyPropertyIsDifferent() {
		SpreadsheetCellStyle one = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("0.00").withAlign(CellStyle.ALIGN_CENTER);
		SpreadsheetCellStyle two = new SpreadsheetCellStyle().withBold(true).withFontName("Courier")
				.withWrapText(true).withDataFormatString("#,##0").withAlign(CellStyle.ALIGN_CENTER);
		// one and two differ only in data format string

		assertThat(one, not(sameInstance(two)));
		assertThat(one.hashCode(), is(not(two.hashCode())));
	}

	@Test
	public void withSurroundBorder_SetsAllFourBordersAtOnce() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withSurroundBorder(CellStyle.BORDER_MEDIUM);

		assertThat(cellStyle.getTopBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(cellStyle.getLeftBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(cellStyle.getBottomBorder(), is(CellStyle.BORDER_MEDIUM));
		assertThat(cellStyle.getRightBorder(), is(CellStyle.BORDER_MEDIUM));
	}

	@Test
	public void CellStyleFont_IsCreatedWhenAFontPropertyIsSet() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle();
		assertThat(cellStyle.getFont(), is(nullValue()));	// no font property set, so font is null (like all properties)

		cellStyle = cellStyle.withAlign(CellStyle.ALIGN_CENTER);	// setting a cell property (rather than a font property)
		assertThat(cellStyle.getFont(), is(nullValue()));			// does not create a font object

		cellStyle = cellStyle.withBold(true);					// setting a font property
		assertThat(cellStyle.getFont(), is(notNullValue()));	// creates an underlying font object
		assertThat(cellStyle.getBold(), is(true));
		assertThat(cellStyle.getFont().getBold(), is(true));
	}

	@Test
	public void settingFontProperties_ChangesTheUnderlyingFont() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withFontOffset(1);	// creates an underlying font object
		SpreadsheetFont originalFont = cellStyle.getFont();
		assertThat(originalFont, is(notNullValue()));
		assertThat(originalFont.getBold(), is(nullValue()));

		cellStyle = cellStyle.withBold(true);
		SpreadsheetFont newFont = cellStyle.getFont();
		assertThat(newFont, not(sameInstance(originalFont)));
		assertThat(newFont.getBold(), is(true));
	}

	@Test
	public void underlineAndDoubleunderline_AreExclusive() {
		SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withUnderline(true);
		assertThat(cellStyle.getUnderline(), is(true));
		assertThat(cellStyle.getDoubleUnderline(), is(nullValue()));

		cellStyle = cellStyle.withDoubleUnderline(true);
		assertThat(cellStyle.getUnderline(), is(nullValue()));
		assertThat(cellStyle.getDoubleUnderline(), is(true));

		cellStyle = cellStyle.withUnderline(true);
		assertThat(cellStyle.getUnderline(), is(true));
		assertThat(cellStyle.getDoubleUnderline(), is(nullValue()));
	}

}
