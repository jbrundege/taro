package taro.spreadsheet;

import org.junit.Test;
import taro.spreadsheet.model.AbstractTest;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.Assert.fail;

public class SpreadsheetReaderTest extends AbstractTest {

    @Test
    public void getColumnIndex_HandlesUpToThreeLetters() {

        assertThat(SpreadsheetReader.getColumnIndex("A1"))
                .isEqualTo(0);

        assertThat(SpreadsheetReader.getColumnIndex("B1"))
                .isEqualTo(1);

        assertThat(SpreadsheetReader.getColumnIndex("Z1"))
                .isEqualTo(25);

        assertThat(SpreadsheetReader.getColumnIndex("AA1"))
                .isEqualTo(26);

        assertThat(SpreadsheetReader.getColumnIndex("AR2"))
                .isEqualTo(43);

        assertThat(SpreadsheetReader.getColumnIndex("BA3"))
                .isEqualTo(52);

        assertThat(SpreadsheetReader.getColumnIndex("ZZ4"))
                .isEqualTo(701);

        assertThat(SpreadsheetReader.getColumnIndex("ABC5"))
                .isEqualTo(730);

        assertThat(SpreadsheetReader.getColumnIndex("ABC6"))
                .isEqualTo(730);

        assertThat(SpreadsheetReader.getColumnIndex("BCL7"))
                .isEqualTo(1441);
    }

    @Test
    public void getColumnIndex_ThrowsExceptionIfColumnMissing() {
        try {
            SpreadsheetReader.getColumnIndex("22");
        } catch(IllegalArgumentException ex) {
            assertThat(ex.getMessage())
                    .contains("22");
        }
    }

    @Test
    public void getColumnIndex_ThrowsExceptionIfCellIdIsMalformed() {
        try {
            SpreadsheetReader.getColumnIndex("B.2");
        } catch(IllegalArgumentException ex) {
            assertThat(ex.getMessage())
                    .contains("cellId");
        }
    }

    @Test
    public void getRowIndex_ReturnsNumberMinusOne() {

        assertThat(SpreadsheetReader.getRowIndex("A1"))
                .isEqualTo(0);

        assertThat(SpreadsheetReader.getRowIndex("Q1"))
                .isEqualTo(0);

        assertThat(SpreadsheetReader.getRowIndex("A2"))
                .isEqualTo(1);

        assertThat(SpreadsheetReader.getRowIndex("ZZZ1492"))
                .isEqualTo(1491);
    }

    @Test
    public void getRowIndex_ThrowsExceptionIfRowMissing() {
        try {
            SpreadsheetReader.getRowIndex("AA");
        } catch(IllegalArgumentException ex) {
            assertThat(ex.getMessage())
                    .contains("AA");
        }
    }

    @Test
    public void getRowIndex_ThrowsExceptionIfCellIdIsMalformed() {
        try {
            SpreadsheetReader.getRowIndex("B2.1");
        } catch(IllegalArgumentException ex) {
            assertThat(ex.getMessage())
                    .contains("2.1");
        }
    }

    @Test
    public void readDown_ReadsStringColumn() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        String[] values = sheet.readDown("A1", 3);

        assertThat(values.length)
                .isEqualTo(3);

        assertThat(values[0])
                .isEqualTo("Fred");

        assertThat(values[1])
                .isEmpty();

        assertThat(values[2])
                .isEqualTo("Sam");
    }

    @Test
    public void readDown_ReadsNumericColumn() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        String[] values = sheet.readDown("B2", 3);

        assertThat(values.length)
                .isEqualTo(3);

        assertThat(values[0])
                .isEqualTo(String.valueOf(-1));

        assertThat(values[1])
                .isEqualTo(String.valueOf(20));

        assertThat(values[2])
                .isEmpty();
    }

    @Test
    public void readDown_ReadsOneColumn() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTab().getPoiSheet());
        String[] values = sheet.readDown("A2", 1);

        assertThat(values.length)
                .isEqualTo(1);

        assertThat(values[0])
                .isEmpty();
    }

    @Test
    public void readDown_ReadsNoColumns() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTab().getPoiSheet());
        String[] values = sheet.readDown("A2", 0);

        assertThat(values.length)
                .isEqualTo(0);
    }

    @Test
    public void readDownNumeric_ReturnsDoubles() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        double[] values = sheet.readDownNumeric("D2", 4);

        assertThat(values.length)
                .isEqualTo(4);

        assertThat(values[0])
                .isEqualTo(4.3d);

        assertThat(values[1])
                .isEqualTo(.154d);

        assertThat(values[2])
                .isEqualTo(-589.1d);

        assertThat(values[3])
                .isEqualTo(0d);    // null is returned as 0!
    }



    @Test
    public void readDownNumeric_ThrowsExceptionIfNonNumericColumn() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        try {
            sheet.readDownNumeric("A1", 1);
            fail("Expected an Exception but not thrown");
        } catch(Exception ex) {
            // expected, this is a weak test anyway, because our mock is throwing the exception, not POI
        }
    }

    @Test
    public void readDownUntilBlank_StopsAtNull() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        List<String> values = sheet.readDownUntilBlank("B2");

        assertThat(values.size())
                .isEqualTo(2);

        assertThat(values.get(0))
                .isEqualTo(String.valueOf(-1));

        assertThat(values.get(1))
                .isEqualTo(String.valueOf(20));
    }

    @Test
    public void readDownUntilBlank_StopsAtEmptyString() {
        SpreadsheetReader sheet = new SpreadsheetReader(getSpreadsheetTabWithValues().getPoiSheet());
        List<String> values = sheet.readDownUntilBlank("A3");

        assertThat(values.size())
                .isEqualTo(2);

        assertThat(values.get(0))
                .isEqualTo("Sam");

        assertThat(values.get(1))
                .isEqualTo("Mary");
    }




}
