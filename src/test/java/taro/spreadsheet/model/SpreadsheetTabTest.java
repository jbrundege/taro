package taro.spreadsheet.model;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;
import static taro.spreadsheet.model.SpreadsheetCellStyle.CENTER;
import static taro.spreadsheet.model.SpreadsheetCellStyle.LEFT;
import static taro.spreadsheet.model.SpreadsheetCellStyle.RIGHT;

public class SpreadsheetTabTest extends AbstractTest {



    @Test
    public void getCellAddress_ConvertsZeroIndexToColRowFormat() {
        assertThat(SpreadsheetTab.getCellAddress(0, 0)
                .equals("A1"));
        assertThat(SpreadsheetTab.getCellAddress(0, 1)
                .equals("B1"));
        assertThat(SpreadsheetTab.getCellAddress(1, 0)
                .equals("A2"));
        assertThat(SpreadsheetTab.getCellAddress(26, 26)
                .equals("AA27"));
        assertThat(SpreadsheetTab.getCellAddress(26, 27)
                .equals("AB27"));
    }

    @Test
    public void setValue_SetsBothValueAndStyle() {
        SpreadsheetTab tab = getSpreadsheetTab();

        tab.setValue(0, 0, "one");

        assertThat(tab.getCell(0, 0).getValue())
                .isEqualTo("one");

        tab.setValue("B2", "two");

        assertThat(tab.getCell("B2").getValue())
                .isEqualTo("two");

        tab.setValue(2, 2, "three", CENTER);
        assertThat(tab.getCell(2, 2).getValue())
                .isEqualTo("three");

        assertThat(tab.getCell(2, 2).getPoiCell().getCellStyle().getAlignmentEnum())
                .isEqualTo(HorizontalAlignment.CENTER);
    }

    @Test
    public void setStyle_SetsStyleOnSingleCell() {
        SpreadsheetTab tab = getSpreadsheetTab();
        tab.setStyle(1, 1, RIGHT);
        assertThat(tab.getCell(1, 1).getStyle())
                .isEqualTo(RIGHT);

        tab.setStyle("B2", CENTER);
        assertThat(tab.getCell(1, 1).getStyle())
                .isEqualTo(CENTER);
    }

    @Test
    public void setStyle_SetsStyleOnBlockOfCells() {
        SpreadsheetTab tab = getSpreadsheetTab();
        tab.setStyle(1, 2, 2, 3, RIGHT);

        // indicated area is styled
        assertThat(tab.getCell(1, 2).getStyle())
                .isEqualTo(RIGHT);
        assertThat(tab.getCell(1, 3).getStyle())
                .isEqualTo(RIGHT);
        assertThat(tab.getCell(2, 2).getStyle())
                .isEqualTo(RIGHT);
        assertThat(tab.getCell(2, 3).getStyle())
                .isEqualTo(RIGHT);

        // surrounding perimeter is null
        assertThat(tab.getOrCreateCell(0, 2).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(0, 3).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(1, 1).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(1, 4).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(2, 1).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(2, 4).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(3, 2).getStyle())
                .isNull();
        assertThat(tab.getOrCreateCell(3, 3).getStyle())
                .isNull();
    }

    @Test
    public void mergeCells_MergesRegion() {
        SpreadsheetTab tab = getSpreadsheetTab();
        tab.mergeCells("B2", "C3", "one big cell", CENTER);

        CellRangeAddress range = tab.getPoiSheet().getMergedRegion(0);

        assertThat(range.getNumberOfCells())
                .isEqualTo(4);
        assertThat(range.getFirstColumn())
                .isEqualTo(1);
        assertThat(range.getLastColumn())
                .isEqualTo(2);
        assertThat(range.getFirstRow())
                .isEqualTo(1);
        assertThat(range.getLastRow())
                .isEqualTo(2);
    }

    @Test
    public void computeRowHeightInPoints_SetsHeightTo1point3TimesFontHeight() {
        int fontSize = 11;
        int numLines = 5;
        double expectedRowHeight = 71.5;    // 11 * 1.3 * 5 rounded to nearest 0.25

        SpreadsheetTab tab = getSpreadsheetTab();

        assertThat((double)tab.computeRowHeightInPoints(fontSize, numLines))
                .isCloseTo(expectedRowHeight, within(0.0000001));
    }

    @Test
    public void computeRowHeightInPoints_DoesNotShrinkRowsFromDefaultHeight() {
        int fontSize = 6;
        int numLines = 1;
        // calculated row height is less than the default height, so is not used
        double calculatedRowHeight = 7.8;    // 6 * 1.3 * 1 rounded to nearest 0.25

        SpreadsheetTab tab = getSpreadsheetTab();

        assertThat((double) tab.computeRowHeightInPoints(fontSize, numLines))
                .isGreaterThan(calculatedRowHeight)
                .isCloseTo(tab.getPoiSheet().getDefaultRowHeightInPoints(), within(0.0000001));
    }

    @Test
    public void autoSizeRow_SetsRowHeightToTallestCell() {
        SpreadsheetTab tab = getSpreadsheetTab();

        tab.setValue(0, 0, "one line");
        tab.setValue(0, 1, "two\nlines");
        tab.setValue(0, 2, "three\nlines\nhere");
        tab.setValue(0, 3, "two\nlines");

        tab.setStyle("A1", "D1", LEFT.withFontSizeInPoints(13));

        double expectedRowHeight = 50.75;     // 13 * 1.3 * 3 rounded to nearest 0.25

        tab.autosizeRows();

        Row row = tab.getPoiSheet().getRow(0);


        assertThat((double) row.getHeightInPoints())
                .isCloseTo(expectedRowHeight, within(0.00000001));
    }

    @Test
    public void setSurroundBorder_AddsBorderToCellRangePerimeter() {
        SpreadsheetTab tab = getSpreadsheetTab();
        tab.setSurroundBorder("B2", "C3", BorderStyle.MEDIUM);

        assertThat(tab.getCell("B2").getStyle().getTopBorder())
                .isEqualTo(BorderStyle.MEDIUM);
        assertThat(tab.getCell("B2").getStyle().getLeftBorder())
                .isEqualTo(BorderStyle.MEDIUM);
        assertThat(tab.getCell("B2").getStyle().getBottomBorder())
                .isNull();

        assertThat(tab.getCell("B2").getStyle().getRightBorder())
                .isNull();

        assertThat(tab.getCell("C2").getStyle().getTopBorder())
                .isEqualTo(BorderStyle.MEDIUM);

        assertThat(tab.getCell("C2").getStyle().getLeftBorder())
                .isNull();
        assertThat(tab.getCell("C2").getStyle().getBottomBorder())
                .isNull();
        assertThat(tab.getCell("C2").getStyle().getRightBorder())
                .isEqualTo(BorderStyle.MEDIUM);

        assertThat(tab.getCell("B3").getStyle().getTopBorder())
                .isNull();
        assertThat(tab.getCell("B3").getStyle().getLeftBorder())
                .isEqualTo(BorderStyle.MEDIUM);
        assertThat(tab.getCell("B3").getStyle().getBottomBorder())
                .isEqualTo(BorderStyle.MEDIUM);
        assertThat(tab.getCell("B3").getStyle().getRightBorder())
                .isNull();

        assertThat(tab.getCell("C3").getStyle().getTopBorder())
                .isNull();
        assertThat(tab.getCell("C3").getStyle().getLeftBorder())
                .isNull();
        assertThat(tab.getCell("C3").getStyle().getBottomBorder())
                .isEqualTo(BorderStyle.MEDIUM);
        assertThat(tab.getCell("C3").getStyle().getRightBorder())
                .isEqualTo(BorderStyle.MEDIUM);
    }

}