package taro.spreadsheet;


import static com.google.common.collect.Lists.newArrayList;
import static org.apache.commons.lang.StringUtils.isNotBlank;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/**
 * Very simple utility to help read a POI sheet within an Excel (.xlsx) file.
 * Does not work with .xls files.
 *
 * The getValue and getStringValue methods return the TRIMMED content of the cell, and return an empty String
 * if the cell doesn't exist or is empty.
 *
 * The getNumericValue method returns 0 if the cell doesn't exist or is empty. It throws an exception if the
 * value cannot be parsed to a number.
 *
 * The getDateValue method returns null if the cell doesn't exist or is empty, and throws an exception if the
 * value is not a date.
 */
public class SpreadsheetReader {

	/**
	 * Gets the 0-based index (used by POI) of the column given a cellId in Excel notation, which must be
	 * one or more letters (case-insensitive) followed by an integer.
	 * i.e. B7 (which returns 1) or AR1677 (which returns 43).
	 *
	 * Throws an IllegalArgumentException if the cellId is malformed
	 */
	public static int getColumnIndex(String cellId) {
		return new CellReference(cellId).getCol();
	}

	/**
	 * Gets the 0-based index (used by POI) of the row given a cellId in Excel notation, which must be
	 * one or more letters (case-insensitive) followed by an integer.
	 * i.e. B7 (which returns 6) or BR1677 (which returns 1676)
	 *
	 * Throws an IllegalArgumentException if the cellId is malformed
	 */
	public static int getRowIndex(String cellId) {
		return new CellReference(cellId).getRow();
	}

	public static String getCellAddress(int col, int row) {
		return CellReference.convertNumToColString(col) + (row+1);
	}


	private Sheet sheet;
	
	public SpreadsheetReader(Sheet sheet) {
		this.sheet = sheet;
	}

	/**
	 * Attempts to convert all values to a string. Returns the trimmed content of the cell, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getValue(String cellId) {
		Cell cell = getCell(cellId);
		return getValue(cell);
	}

	/**
	 * Attempts to convert all values to a string. Returns the trimmed content of the cell, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getValue(int colIndex, int rowIndex) {
		Cell cell = getCell(colIndex, rowIndex);
		return getValue(cell);
	}

	/**
	 * Attempts to convert all values to a string. Returns the trimmed content of the cell, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getValue(Cell cell) {
		if (cell == null) return "";
		
		String value;
		try {
			value = cell.getStringCellValue();
			if (value == null) {
				value = "";
			} else {
				value = value.trim();
			}
		} catch(Exception ex) {
			try {
				value = Double.toString(cell.getNumericCellValue());
			} catch (Exception ex2) {
				throw new RuntimeException("Cannot get value for cell at colIndex: " + cell.getColumnIndex() 
						+ ", rowIndex: " + cell.getRowIndex(), ex);
			}
		}
		return value;
	}

	/**
	 * Returns the trimmed content of the cell as a String, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getStringValue(String cellId) {
		Cell cell = getCell(cellId);
		return getStringValue(cell);
	}

	/**
	 * Returns the trimmed content of the cell as a String, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getStringValue(int columnIndex, int rowIndex) {
		Cell cell = getCell(columnIndex, rowIndex);
		return getStringValue(cell);
	}

	/**
	 * Returns the trimmed content of the cell as a String, or an empty String
	 * if the cell doesn't exist or is empty.
	 */
	public String getStringValue(Cell cell) {
		if (cell == null) {
			return "";
		} else {
			String value = cell.getStringCellValue();
			if (value != null) {
				value = value.trim();
			}
			return value;
		}
	}

	/**
	 * Returns the numeric content of the cell, or 0 if the cell doesn't exist or is empty.
	 */
	public Double getNumericValue(String cellId) {
		Cell cell = getCell(cellId);
		if (cell == null) {
			return 0d;
		} else {
			return cell.getNumericCellValue();
		}
	}

	/**
	 * Returns the numeric content of the cell, or 0 if the cell doesn't exist or is empty.
	 */
	public Double getNumericValue(int columnIndex, int rowIndex) {
		Cell cell = getCell(columnIndex, rowIndex);
		if (cell == null) {
			return 0d;
		} else {
			return cell.getNumericCellValue();
		}
	}

	/**
	 * Returns the Date content of the cell, or null if the cell doesn't exist or is empty.
	 */
	public Date getDateValue(String cellId) {
		Cell cell = getCell(cellId);
		if (cell == null) {
			return null;
		} else {
			return cell.getDateCellValue();
		}
	}

	public Cell getCell(String cellId) {
		int columnIndex = getColumnIndex(cellId);
		int rowIndex = getRowIndex(cellId);
		
		return getCell(columnIndex, rowIndex);
	}
	
	public Cell getCell(int columnIndex, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			return null;
		} else {
			return row.getCell(columnIndex);
		}
	}

	public int getCellType(int col, int row) {
		Cell cell = getCell(col, row);
		return cell != null ? cell.getCellType() : Cell.CELL_TYPE_BLANK;
	}

	public boolean isString(int col, int row) {
		return getCellType(col, row) == Cell.CELL_TYPE_STRING;
	}
	
	public boolean isNumeric(int col, int row) {
		return getCellType(col, row) == Cell.CELL_TYPE_NUMERIC;
	}

	public List<String> readDownUntilBlank(String startingCell) {
		List<String> values = newArrayList(); 
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		String value = getValue(colIndex, rowIndex);
		while (isNotBlank(value)) {
			values.add(value);
			rowIndex++;
			value = getValue(colIndex, rowIndex);
		}
		return values;
	}

	public String[] readDown(String startingCell, int num) {
		String[] values = new String[num]; 
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		for (int i = 0; i < values.length; i++) {
			values[i] = getValue(colIndex, rowIndex+i);
		}
		return values;
	}

	public double[] readDownNumeric(String startingCell, int num) {
		double[] values = new double[num]; 
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		for (int i = 0; i < values.length; i++) {
			values[i] = getNumericValue(colIndex, rowIndex+i);
		}
		return values;
	}

	public List<String> readAcrossUntilBlank(String startingCell) {
		List<String> values = newArrayList();
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		String value = getValue(colIndex, rowIndex);
		while (isNotBlank(value)) {
			values.add(value);
			colIndex++;
			value = getValue(colIndex, rowIndex);
		}
		return values;
	}

	public String[] readAcross(String startingCell, int num) {
		String[] values = new String[num];
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		for (int i = 0; i < values.length; i++) {
			values[i] = getValue(colIndex+i, rowIndex);
		}
		return values;
	}

	public double[] readAcrossNumeric(String startingCell, int num) {
		double[] values = new double[num];
		int rowIndex = getRowIndex(startingCell);
		int colIndex = getColumnIndex(startingCell);
		for (int i = 0; i < values.length; i++) {
			values[i] = getNumericValue(colIndex+i, rowIndex);
		}
		return values;
	}

	public String[][] readSheet() {
		List<List<String>> contents = newArrayList();
		int maxColNum = 0;
		Iterator<Row> rowIterator = sheet.rowIterator();
		while (rowIterator.hasNext()) {
			List<String> rowContents = newArrayList();
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				rowContents.add(getValue(cellIterator.next()));
			}
			contents.add(rowContents);
			if (rowContents.size() > maxColNum) {
				maxColNum = rowContents.size();
			}
		}
		String[][] contentsArray = new String[contents.size()][maxColNum];
		for (int i = 0; i < contentsArray.length; i++) {
			contentsArray[i] = contents.get(i).toArray(new String[maxColNum]);
		}
		return contentsArray;
	}

	public String getSheetName() {
		return sheet.getSheetName();
	}

	public boolean rowHasData(int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		return row != null && row.getPhysicalNumberOfCells() > 0;
	}
}
