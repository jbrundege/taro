package taro.spreadsheet.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Test;

import java.awt.*;
import java.util.Calendar;
import java.util.Date;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;


public class SpreadsheetCellTest extends AbstractTest {



    @Test
    public void setValueWithString_SetsAStringValueOnTheCell() {
        SpreadsheetCell cell = getCell();
        cell.setValue("A String");

        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_STRING);

        assertThat(cell.getPoiCell().getStringCellValue())
                .isEqualTo("A String");
    }

    @Test
    public void setValueWithStringFormula_SetsAFormulaOnTheCell() {
        SpreadsheetCell cell = getCell();

        cell.setValue("=B1*C1");    // formula is any string starting with an equals (=) sign

        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_FORMULA);

        assertThat(cell.getPoiCell().getCellFormula())
                .isEqualTo("B1*C1");
    }

    @Test
    public void setValueWithShort_SetsANumericValueOnTheCell() {
        SpreadsheetCell cell = getCell();
        cell.setValue((short) 12);
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isEqualTo(12.0);
    }

    @Test
    public void setValueWithInteger_SetsANumericValueOnTheCell() {
        SpreadsheetCell cell = getCell();
        cell.setValue(12);
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isEqualTo(12.0);
    }

    @Test
    public void setValueWithLong_SetsANumericValueOnTheCell() {
        SpreadsheetCell cell = getCell();

        cell.setValue(12L);
        
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isEqualTo(12.0);
    }

    @Test
    public void setValueWithFloat_SetsANumericValueOnTheCell() {

        SpreadsheetCell cell = getCell();
        cell.setValue(12f);
        
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isEqualTo(12.0);
    }

    @Test
    public void setValueWithDouble_SetsANumericValueOnTheCell() {
        SpreadsheetCell cell = getCell();
        cell.setValue(12.0);
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isEqualTo(12.0);
    }

    @Test
    public void setValueWithDate_SetsADateValueOnTheCell() {
        Date date = new Date();
        double excelDateNumber = DateUtil.getExcelDate(date);

        SpreadsheetCell cell = getCell();
        cell.setValue(date);
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);
        assertThat(cell.getPoiCell().getNumericCellValue())
                .isCloseTo(excelDateNumber, within(0.001));
    }

    @Test
    public void setValueWithCalendar_SetsADateValueOnTheCell() {
        Calendar calendar = Calendar.getInstance();
        double excelDateNumber = DateUtil.getExcelDate(calendar.getTime());

        SpreadsheetCell cell = getCell();
        cell.setValue(calendar);

        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_NUMERIC);

        assertThat(cell.getPoiCell().getNumericCellValue())
                .isCloseTo(excelDateNumber, within(0.001));
    }

    @Test
    public void setValueWithBoolean_SetsABooleanValueOnTheCell() {
        SpreadsheetCell cell = getCell();
        cell.setValue(true);
        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_BOOLEAN);

        assertThat(cell.getPoiCell().getBooleanCellValue())
                .isTrue();
    }

    @Test
    public void setValueWithNull_SetsABlankValueOnTheCell() {

        SpreadsheetCell cell = getCell();
        cell.setValue("A String to value to start with");
        cell.setValue(null);    // wipe out the string value

        assertThat(cell.getPoiCell().getCellType())
                .isEqualTo(Cell.CELL_TYPE_BLANK);

        assertThat(cell.getPoiCell().getStringCellValue())
                .isEmpty();
    }

    @Test
    public void cellsStyle_IsNullUntilSet() {
        SpreadsheetCell cell = getCell();
        
        assertThat(cell.getStyle())
                .isNull();

        SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withAlign(CellStyle.ALIGN_CENTER);
        cell.setStyle(cellStyle);

        assertThat(cell.getStyle())
                .isNotNull();

        assertThat(cell.getStyle())
                .isEqualTo(cellStyle);
    }

    @Test
    public void applyStyle_MergesAStyleOntoACell() {
        SpreadsheetCellStyle originalStyle = new SpreadsheetCellStyle()
                .withAlign(CellStyle.ALIGN_CENTER)
                .withBold(true)
                .withTopBorder(CellStyle.BORDER_MEDIUM)
                .withLeftBorder(CellStyle.BORDER_MEDIUM)
                .withBottomBorder(CellStyle.BORDER_MEDIUM)
                .withRightBorder(CellStyle.BORDER_MEDIUM);


        SpreadsheetCell cell = getCell();
        cell.setStyle(originalStyle);

        SpreadsheetCellStyle styleToApply = new SpreadsheetCellStyle()
                .withTopBorderColor(Color.RED)
                .withLeftBorderColor(Color.RED)
                .withBottomBorderColor(Color.RED)
                .withRightBorderColor(Color.RED);

        // Method Under Test
        cell.applyStyle(styleToApply);

        SpreadsheetCellStyle currentStyle = cell.getStyle();
        assertThat(currentStyle.getAlign())
                .isEqualTo(CellStyle.ALIGN_CENTER);

        assertThat(currentStyle.getBold())
                .isTrue();

        assertThat(currentStyle.getTopBorder())
                .isEqualTo(CellStyle.BORDER_MEDIUM);

        assertThat(currentStyle.getLeftBorder())
                .isEqualTo(CellStyle.BORDER_MEDIUM);

        assertThat(currentStyle.getBottomBorder())
                .isEqualTo(CellStyle.BORDER_MEDIUM);

        assertThat(currentStyle.getRightBorder())
                .isEqualTo(CellStyle.BORDER_MEDIUM);

        assertThat(currentStyle.getTopBorderColor())
                .isEqualTo(Color.RED);

        assertThat(currentStyle.getLeftBorderColor())
                .isEqualTo(Color.RED);

        assertThat(currentStyle.getBottomBorderColor())
                .isEqualTo(Color.RED);

        assertThat(currentStyle.getRightBorderColor())
                .isEqualTo(Color.RED);
    }

    @Test
    public void getFontSizeInPoints_ReturnsFontSizeIfAStyleWithAFontExists() {
        SpreadsheetCellStyle cellStyle = new SpreadsheetCellStyle().withFontSizeInPoints(15);

        SpreadsheetCell cell = getCell();
        cell.setStyle(cellStyle);
        assertThat(cell.getFontSizeInPoints()).isEqualTo(15);
    }

    @Test
    public void getFontSizeInPoints_ReturnsTheDefaultFontSizeIfStyleOrFontIsNull() {

        SpreadsheetCell cell = getCell();
        
        assertThat(cell.getStyle())
                .isNull();
        assertThat(cell.getFontSizeInPoints())
                .isEqualTo((int) XSSFFont.DEFAULT_FONT_SIZE);

        cell.setStyle(new SpreadsheetCellStyle());

        assertThat(cell.getStyle())
                .isNotNull();
        assertThat(cell.getStyle().getFont())
                .isNull();

        assertThat(cell.getFontSizeInPoints())
                .isEqualTo((int) XSSFFont.DEFAULT_FONT_SIZE);

        cell.setStyle(new SpreadsheetCellStyle().withFontSizeInPoints(17));

        assertThat(cell.getStyle())
                .isNotNull();

        assertThat(cell.getStyle().getFont())
                .isNotNull();

        assertThat(cell.getFontSizeInPoints())
                .isNotEqualTo((int) XSSFFont.DEFAULT_FONT_SIZE);

        assertThat(cell.getFontSizeInPoints())
                .isEqualTo(17);
    }

}
