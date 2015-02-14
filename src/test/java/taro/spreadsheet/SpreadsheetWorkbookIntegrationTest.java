package taro.spreadsheet;


import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import taro.spreadsheet.model.SpreadsheetCellStyle;
import taro.spreadsheet.model.SpreadsheetTab;
import taro.spreadsheet.model.SpreadsheetWorkbook;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;

public class SpreadsheetWorkbookIntegrationTest {

    @Test
    public void createAndVerifyExcelFile() throws IOException {
        // Create a spreadsheet with one tab
        SpreadsheetWorkbook workbook = new SpreadsheetWorkbook();
        SpreadsheetTab tab = workbook.createTab("Test Tab");
        tab.setValue("A1", "Some text", SpreadsheetCellStyle.HEADER.withBackgroundColor(Color.RED));
        tab.setValue("A2", "Some subtext");
        tab.setValue(0, 1, "A multi-line \n text cell", SpreadsheetCellStyle.DEFAULT.withWrapText(true));    // B1
        tab.setValue(1, 1, 27.5, SpreadsheetCellStyle.CENTER_ONE_DECIMAL.withBottomBorder(CellStyle.BORDER_MEDIUM));    // B2
        Date date = new Date();
        tab.setValue("C1", date);

        tab.autosizeRowsAndCols();

        // Write the spreadsheet out as an excel file (to an in-memory byte[])
        byte[] excelFileBytes = writeExcelFileBytes(workbook);
        SpreadsheetReader reader = getReader(excelFileBytes);

        // Verify the spreadsheet (can't verify the styling automatically, but at least verify the text)
        assertThat(reader.isString(0, 0))
                .isTrue();

        assertThat(reader.getStringValue(0, 0))
                .isEqualTo("Some text");    // A1

        assertThat(reader.isString(0, 1))
                .isTrue();

        assertThat(reader.getStringValue(0, 1))
                .isEqualTo("Some subtext");    // A2

        assertThat(reader.isString(1, 0))
                .isTrue();

        assertThat(reader.getStringValue("B1"))
                .isEqualTo("A multi-line \n text cell");

        assertThat(reader.isNumeric(1, 1))
                .isTrue();

        assertThat(reader.getNumericValue("B2"))
                .isCloseTo(27.5, within(0.001));

        assertThat(reader.getDateValue("C1"))
                .isEqualTo(date);
    }

    private byte[] writeExcelFileBytes(SpreadsheetWorkbook workbook) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return out.toByteArray();
    }

    private SpreadsheetReader getReader(byte[] excelFileBytes) throws IOException {
        ByteArrayInputStream in = new ByteArrayInputStream(excelFileBytes);
        XSSFWorkbook poiWorkbook = new XSSFWorkbook(in);
        return new SpreadsheetReader(poiWorkbook.getSheetAt(0));
    }

}
