package taro.spreadsheet.model;

import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
import static org.junit.Assert.*;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Before;
import org.junit.Test;

import static org.assertj.core.api.Assertions.assertThat;

public class SpreadsheetWorkbookTest extends AbstractTest {


    @Before
    public void setup() {
        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();

        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles)
                .isNotNull();

        assertThat(styles)
                .isEmpty();

        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();

        assertThat(fonts)
                .isNotNull();

        assertThat(fonts)
                .isEmpty();
    }

    @Test
    public void registerStyle_CreatesNewStyle_IfNotAlreadyRegistered() {
        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();

        SpreadsheetCellStyle taroStyleOne = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER).withWrapText(true);
        workbook.registerStyle(taroStyleOne);
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();

        assertThat(styles.size())
                .isEqualTo(1);

        CellStyle poiStyleOne = styles.get(taroStyleOne);

        assertThat(poiStyleOne)
                .isNotNull();

        SpreadsheetCellStyle taroStyleTwo = new SpreadsheetCellStyle().withAlign(ALIGN_LEFT).withWrapText(false);
        workbook.registerStyle(taroStyleTwo);
        styles = workbook.getCellStyles();

        assertThat(styles.size())
                .isEqualTo(2);

        CellStyle poiStyleTwo = styles.get(taroStyleTwo);

        assertThat(poiStyleTwo)
                .isNotNull();

        assertThat(poiStyleOne)
                .isNotSameAs(poiStyleTwo);
    }

    @Test
    public void registerStyle_ReusesExistingStyle_IfAlreadyRegistered() {
        SpreadsheetCellStyle styleOne = new SpreadsheetCellStyle()
                .withAlign(ALIGN_CENTER)
                .withWrapText(true)
                .withBold(true);

        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        workbook.registerStyle(styleOne);

        // register one style adds one style
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();

        assertThat(styles.size())
                .isEqualTo(1);

        CellStyle poiStyle = styles.get(styleOne);

        assertThat(poiStyle)
                .isNotNull();

        // register an identiacal style, no new style is added
        SpreadsheetCellStyle styleTwo = new SpreadsheetCellStyle()
                .withAlign(ALIGN_CENTER)
                .withWrapText(true)
                .withBold(true);

        workbook.registerStyle(styleTwo);

        styles = workbook.getCellStyles();

        assertThat(styles.size())
                .isEqualTo(1);

        CellStyle poiStyleOne = styles.get(styleOne);

        assertThat(poiStyleOne)
                .isNotNull();

        CellStyle poiStyleTwo = styles.get(styleTwo);

        assertThat(poiStyleTwo)
                .isNotNull();

        assertThat(poiStyleOne)
                .isSameAs(poiStyleTwo);

        // Font was reused as well
        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();

        assertThat(fonts.size())
                .isEqualTo(1);

        Font poiFontOne = fonts.get(styleOne.getFont());

        assertThat(poiFontOne)
                .isNotNull();

        Font poiFontTwo = fonts.get(styleTwo.getFont());

        assertThat(poiFontTwo)
                .isNotNull();

        assertThat(poiFontOne)
                .isSameAs(poiFontTwo);
    }

    @Test
    public void registerStyle_ReusesFontsIndependentOfStyle() {
        SpreadsheetCellStyle styleOneWithBoldFont = new SpreadsheetCellStyle()
                .withAlign(ALIGN_CENTER)
                .withBold(true);

        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        workbook.registerStyle(styleOneWithBoldFont);

        SpreadsheetCellStyle styleTwoWithBoldFont = new SpreadsheetCellStyle()
                .withAlign(ALIGN_LEFT)
                .withBold(true);

        workbook.registerStyle(styleTwoWithBoldFont);

        // Two independent styles
        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size())
                .isEqualTo(2);

        CellStyle poiStyleOne = styles.get(styleOneWithBoldFont);
        assertThat(poiStyleOne)
                .isNotNull();
        CellStyle poiStyleTwo = styles.get(styleTwoWithBoldFont);
        assertThat(poiStyleTwo).isNotNull();

        assertThat(poiStyleOne)
                .isNotSameAs(poiStyleTwo);

        // One font that was reused
        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();
        assertThat(fonts.size())
                .isEqualTo(1);

        Font poiFontOne = fonts.get(styleOneWithBoldFont.getFont());
        assertThat(poiFontOne).isNotNull();
        Font poiFontTwo = fonts.get(styleTwoWithBoldFont.getFont());
        assertThat(poiFontTwo).isNotNull();

        assertThat(poiFontOne)
                .isSameAs(poiFontTwo);
    }

    @Test
    public void getStyles_ReturnsImmutableMap() {
        SpreadsheetCellStyle style = new SpreadsheetCellStyle().withAlign(ALIGN_CENTER).withBold(true);

        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        workbook.registerStyle(style);

        Map<SpreadsheetCellStyle, CellStyle> styles = workbook.getCellStyles();
        assertThat(styles.size())
                .isEqualTo(1);
        assertThat(styles.keySet())
                .contains(style);
        try {
            styles.put(style, null);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }

        try {
            styles.clear();
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }

        try {
            styles.remove(style);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }
    }

    @Test
    public void getFonts_ReturnsImmutableMap() {
        SpreadsheetCellStyle style = new SpreadsheetCellStyle().withBold(true);
        SpreadsheetFont font = style.getFont();

        assertThat(font)
                .isNotNull();

        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        workbook.registerStyle(style);

        Map<SpreadsheetFont, Font> fonts = workbook.getFonts();

        assertThat(fonts.size())
                .isEqualTo(1);

        assertThat(fonts.keySet())
                .contains(font);

        try {
            fonts.put(font, null);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }

        try {
            fonts.clear();
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }

        try {
            fonts.remove(font);
            fail("Expected an UnsupportedOperationException but not thrown.");
        } catch (UnsupportedOperationException ex) { /* expected */ }
    }

    @Test
    public void createTab_CachesByIndexAndTitle() {
        String title0 = "tab at index 0";
        String title1 = "tab at index 1";
        String title2 = "tab at index 2";

        SpreadsheetWorkbook workbook = getSpreadsheetWorkbook();
        SpreadsheetTab tab0 = workbook.createTab(title0);
        SpreadsheetTab tab1 = workbook.createTab(title1);
        SpreadsheetTab tab2 = workbook.createTab(title2);

        assertThat(workbook.getTab(0))
                .isSameAs(tab0);

        assertThat(workbook.getTab(title0))
                .isSameAs(tab0);

        assertThat(workbook.getTab(1))
                .isSameAs(tab1);

        assertThat(workbook.getTab(title1))
                .isSameAs(tab1);

        assertThat(workbook.getTab(2))
                .isSameAs(tab2);

        assertThat(workbook.getTab(title2))
                .isSameAs(tab2);
    }

}
