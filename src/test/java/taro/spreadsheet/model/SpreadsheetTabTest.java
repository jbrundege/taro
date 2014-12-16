package taro.spreadsheet.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.ArgumentCaptor;
import org.mockito.Mock;

import static org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.closeTo;
import static org.mockito.Matchers.anyInt;
import static org.mockito.Mockito.*;
import static org.mockito.MockitoAnnotations.initMocks;
import static taro.spreadsheet.model.SpreadsheetCellStyle.CENTER;
import static taro.spreadsheet.model.SpreadsheetCellStyle.LEFT;
import static taro.spreadsheet.model.SpreadsheetCellStyle.RIGHT;

public class SpreadsheetTabTest {

	private SpreadsheetWorkbook workbook;
	private SpreadsheetTab tab;

	@Mock private Sheet poiSheet;
	@Mock private Row poiRow;
	@Mock private Cell poiCell;

	@Before
	public void setup() {
		initMocks(this);
		workbook = new SpreadsheetWorkbook();
		when(poiSheet.getRow(anyInt())).thenReturn(poiRow);
		when(poiRow.getCell(anyInt())).thenReturn(poiCell);
		tab = new SpreadsheetTab(workbook, poiSheet);
	}

	@Test
	public void getCellAddress_ConvertsZeroIndexToColRowFormat() {
		assertThat(SpreadsheetTab.getCellAddress(0, 0), is("A1"));
		assertThat(SpreadsheetTab.getCellAddress(0, 1), is("B1"));
		assertThat(SpreadsheetTab.getCellAddress(1, 0), is("A2"));
		assertThat(SpreadsheetTab.getCellAddress(26, 26), is("AA27"));
		assertThat(SpreadsheetTab.getCellAddress(26, 27), is("AB27"));
	}

	@Test
	public void constructor_CreatesSheetFromTitle() {
		XSSFWorkbook poiWorkbook = mock(XSSFWorkbook.class);
		workbook = new SpreadsheetWorkbook(poiWorkbook);
		new SpreadsheetTab(workbook, "testing");
		verify(poiWorkbook).createSheet("testing");
	}

	@Test
	public void setValue_SetsBothValueAndStyle() {
		ArgumentCaptor<String> stringCaptor = ArgumentCaptor.forClass(String.class);

		tab.setValue(0, 0, "one");
		verify(poiCell).setCellValue(stringCaptor.capture());
		assertThat(stringCaptor.getValue(), is("one"));
		reset(poiCell);

		tab.setValue("B2", "two");
		verify(poiCell).setCellValue(stringCaptor.capture());
		assertThat(stringCaptor.getValue(), is("two"));
		reset(poiCell);

		ArgumentCaptor<CellStyle> cellStyleCaptor = ArgumentCaptor.forClass(CellStyle.class);

		tab.setValue(2, 2, "three", CENTER);
		verify(poiCell).setCellValue(stringCaptor.capture());
		assertThat(stringCaptor.getValue(), is("three"));
		verify(poiCell).setCellStyle(cellStyleCaptor.capture());
		CellStyle style = cellStyleCaptor.getValue();
		assertThat(style.getAlignment(), is(CellStyle.ALIGN_CENTER));
	}

	@Test
	public void setStyle_SetsStyleOnSingleCell() {
		tab.setStyle(1, 1, RIGHT);
		assertThat(tab.getCell(1, 1).getStyle(), is(RIGHT));

		tab.setStyle("B2", CENTER);
		assertThat(tab.getCell(1, 1).getStyle(), is(CENTER));
	}

	@Test
	public void setStyle_SetsStyleOnBlockOfCells() {
		tab.setStyle(1, 2, 2, 3, RIGHT);

		// indicated area is styled
		assertThat(tab.getCell(1, 2).getStyle(), is(RIGHT));
		assertThat(tab.getCell(1, 3).getStyle(), is(RIGHT));
		assertThat(tab.getCell(2, 2).getStyle(), is(RIGHT));
		assertThat(tab.getCell(2, 3).getStyle(), is(RIGHT));

		// surrounding perimeter is null
		assertThat(tab.getCell(0, 2).getStyle(), is(nullValue()));
		assertThat(tab.getCell(0, 3).getStyle(), is(nullValue()));
		assertThat(tab.getCell(1, 1).getStyle(), is(nullValue()));
		assertThat(tab.getCell(1, 4).getStyle(), is(nullValue()));
		assertThat(tab.getCell(2, 1).getStyle(), is(nullValue()));
		assertThat(tab.getCell(2, 4).getStyle(), is(nullValue()));
		assertThat(tab.getCell(3, 2).getStyle(), is(nullValue()));
		assertThat(tab.getCell(3, 3).getStyle(), is(nullValue()));
	}

	@Test
	public void mergeCells_MergesRegion() {
		tab.mergeCells("B2", "C3", "one big cell", CENTER);

		ArgumentCaptor<CellRangeAddress> rangeCaptor = ArgumentCaptor.forClass(CellRangeAddress.class);
		verify(poiSheet).addMergedRegion(rangeCaptor.capture());
		CellRangeAddress range = rangeCaptor.getValue();
		assertThat(range.getNumberOfCells(), is(4));
		assertThat(range.getFirstColumn(), is(1));
		assertThat(range.getLastColumn(), is(2));
		assertThat(range.getFirstRow(), is(1));
		assertThat(range.getLastRow(), is(2));
	}

	@Test
	public void computeRowHeightInPoints_SetsHeightTo1point3TimesFontHeight() {
		int fontSize = 11;
		int numLines = 5;
		double expectedRowHeight = 71.5;	// 11 * 1.3 * 5 rounded to nearest 0.25

		assertThat((double)tab.computeRowHeightInPoints(fontSize, numLines), closeTo(expectedRowHeight, 0.0000001));
	}

	@Test
	public void autoSizeRow_SetsRowHeightToTallestCell() {
		setCellValue(0, 0, "one line");
		setCellValue(0, 1, "two\nlines");
		setCellValue(0, 2, "three\nlines\nhere");
		setCellValue(0, 3, "two\nlines");

		tab.setStyle("A1", "D1", LEFT.withFontSizeInPoints(13));

		double expetedRowHeight = 50.75; 	// 13 * 1.3 * 3 rounded to nearest 0.25

		tab.autosizeRows();

		ArgumentCaptor<Float> floatCaptor = ArgumentCaptor.forClass(Float.class);
		verify(poiRow).setHeightInPoints(floatCaptor.capture());
		assertThat((double)floatCaptor.getValue(), closeTo(expetedRowHeight, 0.00000001));
	}

	@Test
	public void setSurroundBorder_AddsBorderToCellRangePerimeter() {
		tab.setSurroundBorder("B2", "C3", BORDER_MEDIUM);

		assertThat(tab.getCell("B2").getStyle().getTopBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("B2").getStyle().getLeftBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("B2").getStyle().getBottomBorder(), is(nullValue()));
		assertThat(tab.getCell("B2").getStyle().getRightBorder(), is(nullValue()));

		assertThat(tab.getCell("C2").getStyle().getTopBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("C2").getStyle().getLeftBorder(), is(nullValue()));
		assertThat(tab.getCell("C2").getStyle().getBottomBorder(), is(nullValue()));
		assertThat(tab.getCell("C2").getStyle().getRightBorder(), is(BORDER_MEDIUM));

		assertThat(tab.getCell("B3").getStyle().getTopBorder(), is(nullValue()));
		assertThat(tab.getCell("B3").getStyle().getLeftBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("B3").getStyle().getBottomBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("B3").getStyle().getRightBorder(), is(nullValue()));

		assertThat(tab.getCell("C3").getStyle().getTopBorder(), is(nullValue()));
		assertThat(tab.getCell("C3").getStyle().getLeftBorder(), is(nullValue()));
		assertThat(tab.getCell("C3").getStyle().getBottomBorder(), is(BORDER_MEDIUM));
		assertThat(tab.getCell("C3").getStyle().getRightBorder(), is(BORDER_MEDIUM));
	}

	private void setCellValue(int row, int col, String value) {
		Cell cell = mock(Cell.class);
		when(cell.getCellType()).thenReturn(Cell.CELL_TYPE_STRING);
		when(cell.getStringCellValue()).thenReturn(value);
		when(poiRow.getCell(col)).thenReturn(cell);
		tab.setValue(row, col, value);
	}

}