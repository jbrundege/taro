package taro.spreadsheet.model;

import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

public class SpreadsheetWorkbookTest {

    private SpreadsheetWorkbook workbook;

    @Before
    public void setup() {
        workbook = new SpreadsheetWorkbook();

        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles, notNullValue());
        assertThat(styles.isEmpty(), is(true));

        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();
        assertThat(fonts, notNullValue());
        assertThat(fonts.isEmpty(), is(true));
    }

    @Test
    public void registerStyle_CreatesNewStyle_IfNotAlreadyRegistered() {
        SpreadsheetCellStyle taroStyleOne = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER).withWrapText(true);
        workbook.registerStyle(taroStyleOne);
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size(), is(1));
        CellStyle poiStyleOne = styles.get(taroStyleOne);
        assertThat(poiStyleOne, notNullValue());

        SpreadsheetCellStyle taroStyleTwo = new SpreadsheetCellStyle().withAlign(ALIGN_LEFT).withWrapText(false);
        workbook.registerStyle(taroStyleTwo);
        styles = workbook.getCellStyles();
        assertThat(styles.size(), is(2));
        CellStyle poiStyleTwo = styles.get(taroStyleTwo);
        assertThat(poiStyleTwo, notNullValue());

        assertThat(poiStyleOne, not(sameInstance(poiStyleTwo)));
    }

    @Test
    public void registerStyle_ReusesExistingStyle_IfAlreadyRegistered() {
        SpreadsheetCellStyle styleOne = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER)
                .withWrapText(true).withBold(true);
        workbook.registerStyle(styleOne);

        // register one style adds one style
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size(), is(1));
        CellStyle poiStyle = styles.get(styleOne);
        assertThat(poiStyle, notNullValue());

        // register an identiacal style, no new style is added
        SpreadsheetCellStyle styleTwo = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER)
                .withWrapText(true).withBold(true);
        workbook.registerStyle(styleTwo);

        styles = workbook.getCellStyles();
        assertThat(styles.size(), is(1));

        CellStyle poiStyleOne = styles.get(styleOne);
        assertThat(poiStyleOne, notNullValue());
        CellStyle poiStyleTwo = styles.get(styleTwo);
        assertThat(poiStyleTwo, notNullValue());

        assertThat(poiStyleOne, sameInstance(poiStyleTwo));

        // Font was reused as well
        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();
        assertThat(fonts.size(), is(1));

        Font poiFontOne = fonts.get(styleOne.getFont());
        assertThat(poiFontOne, notNullValue());
        Font poiFontTwo = fonts.get(styleTwo.getFont());
        assertThat(poiFontTwo, notNullValue());

        assertThat(poiFontOne, sameInstance(poiFontTwo));
    }

    @Test
    public void registerStyle_ReusesFontsIndependentOfStyle() {
        SpreadsheetCellStyle styleOneWithBoldFont = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER).withBold(true);
        workbook.registerStyle(styleOneWithBoldFont);

        SpreadsheetCellStyle styleTwoWithBoldFont = new SpreadsheetCellStyle().withAlign(ALIGN_LEFT).withBold(true);
        workbook.registerStyle(styleTwoWithBoldFont);

        // Two independent styles
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size(), is(2));

        CellStyle poiStyleOne = styles.get(styleOneWithBoldFont);
        assertThat(poiStyleOne, notNullValue());
        CellStyle poiStyleTwo = styles.get(styleTwoWithBoldFont);
        assertThat(poiStyleTwo, notNullValue());

        assertThat(poiStyleOne, not(sameInstance(poiStyleTwo)));

        // One font that was reused
        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();
        assertThat(fonts.size(), is(1));

        Font poiFontOne = fonts.get(styleOneWithBoldFont.getFont());
        assertThat(poiFontOne, notNullValue());
        Font poiFontTwo = fonts.get(styleTwoWithBoldFont.getFont());
        assertThat(poiFontTwo, notNullValue());

        assertThat(poiFontOne, sameInstance(poiFontTwo));
    }

    @Test
    public void getStyles_ReturnsImmutableMap() {
        SpreadsheetCellStyle style = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER).withBold(true);
        workbook.registerStyle(style);

        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size(), is(1));
        assertThat(styles.keySet(), contains(style));

        try {
            styles.put(style, null);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }

        try {
            styles.clear();
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }

        try {
            styles.remove(style);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }
    }

    @Test
    public void getFonts_ReturnsImmutableMap() {
        SpreadsheetCellStyle style = new SpreadsheetCellStyle().withBold(true);
        SpreadsheetFont font = style.getFont();
        assertThat(font, notNullValue());
        workbook.registerStyle(style);

        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();
        assertThat(fonts.size(), is(1));
        assertThat(fonts.keySet(), contains(font));

        try {
            fonts.put(font, null);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }

        try {
            fonts.clear();
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }

        try {
            fonts.remove(font);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch(UnsupportedOperationException ex) { /* expected */ }
    }

    @Test
    public void createTab_CachesByIndexAndTitle() {
        String title0 = "tab at index 0";
        String title1 = "tab at index 1";
        String title2 = "tab at index 2";

        SpreadsheetTab tab0 = workbook.createTab(title0);
        SpreadsheetTab tab1 = workbook.createTab(title1);
        SpreadsheetTab tab2 = workbook.createTab(title2);

        Assert.assertThat(workbook.getTab(0), sameInstance(tab0));
        Assert.assertThat(workbook.getTab(title0), sameInstance(tab0));
        Assert.assertThat(workbook.getTab(1), sameInstance(tab1));
        Assert.assertThat(workbook.getTab(title1), sameInstance(tab1));
        Assert.assertThat(workbook.getTab(2), sameInstance(tab2));
        Assert.assertThat(workbook.getTab(title2), sameInstance(tab2));
    }

}
