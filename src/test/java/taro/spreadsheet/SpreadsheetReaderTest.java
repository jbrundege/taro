package taro.spreadsheet;

import static org.hamcrest.Matchers.*;
import static org.junit.Assert.*;
import static org.mockito.Matchers.anyInt;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;
import org.mockito.invocation.InvocationOnMock;
import org.mockito.stubbing.Answer;


public class SpreadsheetReaderTest {

	@Test
	public void getColumnIndex_HandlesUpToThreeLetters() {
		assertThat(SpreadsheetReader.getColumnIndex("A1"), is(0));
		assertThat(SpreadsheetReader.getColumnIndex("B1"), is(1));
		assertThat(SpreadsheetReader.getColumnIndex("Z1"), is(25));
		assertThat(SpreadsheetReader.getColumnIndex("AA1"), is(26));
		assertThat(SpreadsheetReader.getColumnIndex("AR2"), is(43));
		assertThat(SpreadsheetReader.getColumnIndex("BA3"), is(52));
		assertThat(SpreadsheetReader.getColumnIndex("ZZ4"), is(701));
		assertThat(SpreadsheetReader.getColumnIndex("ABC5"), is(730));
		assertThat(SpreadsheetReader.getColumnIndex("ABC6"), is(730));
		assertThat(SpreadsheetReader.getColumnIndex("BCL7"), is(1441));
	}

	@Test
	public void getColumnIndex_ThrowsExceptionIfColumnMissing() {
		try {
			SpreadsheetReader.getColumnIndex("22");
		} catch(IllegalArgumentException ex) {
			assertThat(ex.getMessage(), containsString("22"));
		}
	}

	@Test
	public void getColumnIndex_ThrowsExceptionIfCellIdIsMalformed() {
		try {
			SpreadsheetReader.getColumnIndex("B.2");
		} catch(IllegalArgumentException ex) {
			assertThat(ex.getMessage(), containsString("cellId"));
		}
	}
	
	@Test
	public void getRowIndex_ReturnsNumberMinusOne() {
		assertThat(SpreadsheetReader.getRowIndex("A1"), is(0));
		assertThat(SpreadsheetReader.getRowIndex("Q1"), is(0));
		assertThat(SpreadsheetReader.getRowIndex("A2"), is(1));
		assertThat(SpreadsheetReader.getRowIndex("ZZZ1492"), is(1491));
	}
	
	@Test
	public void getRowIndex_ThrowsExceptionIfRowMissing() {
		try {
			SpreadsheetReader.getRowIndex("AA");
		} catch(IllegalArgumentException ex) {
			assertThat(ex.getMessage(), containsString("AA"));
		}
	}
	
	@Test
	public void getRowIndex_ThrowsExceptionIfCellIdIsMalformed() {
		try {
			SpreadsheetReader.getRowIndex("B2.1");
		} catch(IllegalArgumentException ex) {
			assertThat(ex.getMessage(), containsString("2.1"));
		}
	}
	
	@Test 
	public void readDown_ReadsStringColumn() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		String[] values = sheet.readDown("A1", 3);
		assertThat(values.length, is(3));
		assertThat(values[0], is("Fred"));
		assertThat(values[1], is(""));
		assertThat(values[2], is("Sam"));
	}

	@Test 
	public void readDown_ReadsNumericColumn() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		String[] values = sheet.readDown("B2", 3);
		assertThat(values.length, is(3));
		assertThat(values[0], is(Double.toString(-1d)));
		assertThat(values[1], is(Double.toString(20d)));
		assertThat(values[2], is(""));
	}
	
	@Test 
	public void readDown_ReadsOneColumn() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		String[] values = sheet.readDown("A2", 1);
		assertThat(values.length, is(1));
		assertThat(values[0], is(""));
	}
	
	@Test 
	public void readDown_ReadsNoColumns() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		String[] values = sheet.readDown("A2", 0);
		assertThat(values.length, is(0));
	}
	
	@Test 
	public void readDownNumeric_ReturnsDoubles() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		double[] values = sheet.readDownNumeric("D2", 4);
		assertThat(values.length, is(4));
		assertThat(values[0], is(4.3d));
		assertThat(values[1], is(.154d));
		assertThat(values[2], is(-589.1d));
		assertThat(values[3], is(0d));	// null is returned as 0!
	}
	
	@Test 
	public void readDownNumeric_ThrowsExceptionIfNonNumericColumn() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		try {
			sheet.readDownNumeric("A1", 1);
			fail("Expected an Exception but not thrown");
		} catch(Exception ex) {
			// expected, this is a weak test anyway, because our mock is throwing the exception, not POI
		}
	}
	
	@Test 
	public void readDownUntilBlank_StopsAtNull() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		List<String> values = sheet.readDownUntilBlank("B2");
		assertThat(values.size(), is(2));
		assertThat(values.get(0), is(Double.toString(-1d)));
		assertThat(values.get(1), is(Double.toString(20d)));
	}
	
	@Test 
	public void readDownUntilBlank_StopsAtEmptyString() {
		SpreadsheetReader sheet = new SpreadsheetReader(createMockSheet());
		List<String> values = sheet.readDownUntilBlank("A3");
		assertThat(values.size(), is(2));
		assertThat(values.get(0), is("Sam"));
		assertThat(values.get(1), is("Mary"));
	}
	
	
	private Object[][] mockSpreadsheetData = {
		//			A			B		C		D			E
		/* 1 */	{	"Fred",		7d,		null,	2.7d,		null		},
		/* 2 */	{	null,		-1d,	null,	4.3d,		new Date()	},
		/* 3 */	{	"Sam",		20d,	"OH",	.154d,		new Date()	},
		/* 4 */	{	"Mary",		null,	null,	-589.1d,	new Date()	},
		/* 5 */	{	"",			4d,		"B",	null,		new Date()	}
	};
	
	/**
	 * POI is really hard to test. Uses final classes and concrete class return types (rather than interfaces).
	 * To work around this, the following ugly mockito code just makes a sheet that can return values from the above Object[][].
	 * Doesn't verify anything, just acts like a POI sheet.
	 */
	private Sheet createMockSheet() {
		Sheet mockSheet = mock(Sheet.class);
		when(mockSheet.getRow(anyInt())).then(new Answer<Row>() {
			public Row answer(InvocationOnMock invocation) {
				final Integer rowIndex = (Integer)invocation.getArguments()[0];
				Row mockRow = mock(Row.class);
				when(mockRow.getCell(anyInt())).then(new Answer<Cell>() {
					public Cell answer(InvocationOnMock invocation) {
						Integer colIndex = (Integer)invocation.getArguments()[0];
						Cell mockCell = mock(Cell.class);

						Object value = mockSpreadsheetData[rowIndex][colIndex];
						if (value == null) {
							when(mockCell.getCellType()).thenReturn(Cell.CELL_TYPE_BLANK);
							return mockCell;
						}

						try {
							String retVal = (String)mockSpreadsheetData[rowIndex][colIndex];
							when(mockCell.getStringCellValue()).thenReturn(retVal);
							RichTextString richTextString = mock(RichTextString.class);
							when(richTextString.getString()).thenReturn(retVal);
							when(mockCell.getRichStringCellValue()).thenReturn(richTextString);
							when(mockCell.getCellType()).thenReturn(Cell.CELL_TYPE_STRING);
						} catch(ClassCastException ex) {
							when(mockCell.getStringCellValue()).thenThrow(new IllegalStateException());
						}
						try {
							Date retVal = (Date)mockSpreadsheetData[rowIndex][colIndex];
							when(mockCell.getDateCellValue()).thenReturn(retVal);
							when(mockCell.getCellType()).thenReturn(Cell.CELL_TYPE_NUMERIC);
						} catch(ClassCastException ex) {
							when(mockCell.getDateCellValue()).thenThrow(new IllegalStateException());
						}
						try {
							double doubleVal = 0d;
							if (mockSpreadsheetData[rowIndex][colIndex] != null) {
								doubleVal = (Double)mockSpreadsheetData[rowIndex][colIndex];
							}
							when(mockCell.getNumericCellValue()).thenReturn(doubleVal);
							when(mockCell.getCellType()).thenReturn(Cell.CELL_TYPE_NUMERIC);
						} catch(ClassCastException ex) {
							when(mockCell.getNumericCellValue()).thenThrow(new NumberFormatException());
						}
						return mockCell;
					}
				});
				return mockRow;
			}
		});
		return mockSheet;
	}
	
}
