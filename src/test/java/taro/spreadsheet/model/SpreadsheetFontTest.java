package taro.spreadsheet.model;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.not;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.CoreMatchers.sameInstance;
import static org.hamcrest.MatcherAssert.assertThat;

import org.junit.Test;

public class SpreadsheetFontTest {

    @Test
    public void aNewSpreadsheetFont_HasAllNullProperties() {
        SpreadsheetFont font = new SpreadsheetFont();
        assertThat(font.getBold(), nullValue());
        assertThat(font.getItalic(), nullValue());
        assertThat(font.getStrikeout(), nullValue());
        assertThat(font.getUnderline(), nullValue());
        assertThat(font.getDoubleUnderline(), nullValue());
        assertThat(font.getFontName(), nullValue());
        assertThat(font.getFontOffset(), nullValue());
        assertThat(font.getFontSizeInPoints(), nullValue());
    }

    @Test
    public void spreadsheetFont_IsImmutable() {
        SpreadsheetFont font = new SpreadsheetFont();
        assertThat(font.withBold(true), is(not(font)));
        assertThat(font.withDoubleUnderline(true), is(not(font)));
        assertThat(font.withFontName("Courier"), is(not(font)));
        assertThat(font.withFontOffset(1), is(not(font)));
        assertThat(font.withFontSizeInPoints(14), is(not(font)));
        assertThat(font.withItalic(true), is(not(font)));
        assertThat(font.withStrikeout(true), is(not(font)));
        assertThat(font.withUnderline(true), is(not(font)));

        assertThat(font.copy(), not(sameInstance(font)));
        assertThat(font.apply(new SpreadsheetFont()), not(sameInstance(font)));
    }

    @Test
    public void apply_TransfersAllPropertiesToCopy() {
        SpreadsheetFont src = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        SpreadsheetFont dest = new SpreadsheetFont();

        assertThat(dest.getBold(), nullValue());
        assertThat(dest.getItalic(), nullValue());
        assertThat(dest.getStrikeout(), nullValue());
        assertThat(dest.getUnderline(), nullValue());
        assertThat(dest.getFontName(), nullValue());
        assertThat(dest.getFontOffset(), nullValue());
        assertThat(dest.getFontSizeInPoints(), nullValue());

        // Method Under Test
        SpreadsheetFont applied = dest.apply(src);

        assertThat(applied.getBold(), is(true));
        assertThat(applied.getItalic(), is(true));
        assertThat(applied.getStrikeout(), is(true));
        assertThat(applied.getUnderline(), is(true));
        assertThat(applied.getFontName(), is("Courier"));
        assertThat(applied.getFontOffset(), is(1));
        assertThat(applied.getFontSizeInPoints(), is(14));
    }

    @Test
    public void apply_OverwritesNonNullProperties() {
        SpreadsheetFont src = new SpreadsheetFont().withBold(true).withFontSizeInPoints(14);
        SpreadsheetFont dest = new SpreadsheetFont().withDoubleUnderline(true).withFontSizeInPoints(9);

        assertThat(src.getBold(), is(true));
        assertThat(src.getFontSizeInPoints(), is(14));
        assertThat(src.getDoubleUnderline(), nullValue());

        assertThat(dest.getBold(), nullValue());
        assertThat(dest.getFontSizeInPoints(), is(9));
        assertThat(dest.getDoubleUnderline(), is(true));

        // Method Under Test
        SpreadsheetFont applied = dest.apply(src);

        // src was set, so overwrite the null value on dest
        assertThat(applied.getBold(), is(true));

        // src was set, so overwrite the previously set value on dest
        assertThat(applied.getFontSizeInPoints(), is(14));

        // src was not set, so do not overwrite the existing set value on dest
        assertThat(applied.getDoubleUnderline(), is(true));
    }

    @Test
    public void copy_ReturnsNewInstanceWithSameProperties() {
        SpreadsheetFont original = new SpreadsheetFont().withBold(true).withItalic(true).withFontName("Courier")
                .withFontSizeInPoints(14);

        assertThat(original.getBold(), is(true));
        assertThat(original.getItalic(), is(true));
        assertThat(original.getStrikeout(), nullValue());
        assertThat(original.getUnderline(), nullValue());
        assertThat(original.getFontName(), is("Courier"));
        assertThat(original.getFontOffset(), nullValue());
        assertThat(original.getFontSizeInPoints(), is(14));

        // Method Under Test
        SpreadsheetFont copy = original.copy();

        assertThat(copy, not(sameInstance(original)));

        assertThat(copy.getBold(), is(true));
        assertThat(copy.getItalic(), is(true));
        assertThat(copy.getStrikeout(), nullValue());
        assertThat(copy.getUnderline(), nullValue());
        assertThat(copy.getFontName(), is("Courier"));
        assertThat(copy.getFontOffset(), nullValue());
        assertThat(copy.getFontSizeInPoints(), is(14));
    }

    @Test
    public void equals_IsTrueWhenDifferentFontsHaveTheSameProperties() {
        SpreadsheetFont one = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        SpreadsheetFont two = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);

        assertThat(one, not(sameInstance(two)));
        assertThat(one.equals(two), is(true));
    }

    @Test
    public void equals_IsFalseIfAnyPropertyIsDifferent() {
        SpreadsheetFont one = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        SpreadsheetFont two = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(2).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        // one and two differ only in font offset

        assertThat(one, not(sameInstance(two)));
        assertThat(one.equals(two), is(false));
    }

    @Test
    public void hashCode_IsSameWhenDifferentFontsHaveTheSameProperties() {
        SpreadsheetFont one = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        SpreadsheetFont two = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);

        assertThat(one, not(sameInstance(two)));
        assertThat(one.hashCode(), is(two.hashCode()));
    }

    @Test
    public void hashCode_IsDifferentIfAnyPropertyIsDifferent() {
        SpreadsheetFont one = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(1).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        SpreadsheetFont two = new SpreadsheetFont().withBold(true).withFontName("Courier")
                .withFontOffset(2).withFontSizeInPoints(14).withItalic(true).withStrikeout(true).withUnderline(true);
        // one and two differ only in font offset

        assertThat(one, not(sameInstance(two)));
        assertThat(one.hashCode(), is(not(two.hashCode())));
    }

    @Test
    public void underlineAndDoubleunderline_AreExclusive() {
        SpreadsheetFont font = new SpreadsheetFont().withUnderline(true);
        assertThat(font.getUnderline(), is(true));
        assertThat(font.getDoubleUnderline(), is(nullValue()));

        font = font.withDoubleUnderline(true);
        assertThat(font.getUnderline(), is(nullValue()));
        assertThat(font.getDoubleUnderline(), is(true));

        font = font.withUnderline(true);
        assertThat(font.getUnderline(), is(true));
        assertThat(font.getDoubleUnderline(), is(nullValue()));
    }

}