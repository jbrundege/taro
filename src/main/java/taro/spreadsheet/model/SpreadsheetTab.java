package taro.spreadsheet.model;

import static com.google.common.collect.Maps.*;
import static taro.spreadsheet.model.SpreadsheetCellStyle.DEFAULT;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class SpreadsheetTab {

	private SpreadsheetWorkbook workbook;
	private Sheet sheet;
	private Map<String, SpreadsheetCell> cells = newHashMap();
	private Drawing drawing;

	private int highestModifiedCol = -1;
	private int highestModifiedRow = -1;

	public SpreadsheetTab(SpreadsheetWorkbook workbook, String title) {
		this.workbook = workbook;
		this.sheet = workbook.getPoiWorkbook().createSheet(title);
	}

	public SpreadsheetTab(SpreadsheetWorkbook workbook, Sheet sheet) {
		this.workbook = workbook;
		this.sheet = sheet;
	}

	public static String getCellAddress(int row, int col) {
		return CellReference.convertNumToColString(col) + (row+1);
	}

	public void setValue(String cellAddress, Object content) {
		setValue(cellAddress, content, null);
	}

	public void setValue(String cellAddress, Object content, SpreadsheetCellStyle style) {
		CellReference cellReference = new CellReference(cellAddress);
		setValue(cellReference.getRow(), cellReference.getCol(), content, style);
	}

	public void setValue(int row, int col, Object content) {
		setValue(row, col, content, null);
	}

	public void setValue(int row, int col, Object content, SpreadsheetCellStyle style) {
		SpreadsheetCell cell = getCell(row, col);
		cell.setValue(content);
		if (style != null) {
			cell.setStyle(style);
		}
		recordCellModified(row, col);
	}

	public void setStyle(String cellAddress, SpreadsheetCellStyle style) {
		CellReference cellReference = new CellReference(cellAddress);
		setStyle(cellReference.getRow(), cellReference.getCol(), style);
	}

	public void setStyle(String firstCell, String lastCell, SpreadsheetCellStyle style) {
		CellReference firstReference = new CellReference(firstCell);
		CellReference lastReference = new CellReference(lastCell);
		setStyle(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), style);
	}

	public void setStyle(int row, int col, SpreadsheetCellStyle style) {
		getCell(row, col).setStyle(style);
	}

	public void setStyle(int firstRow, int lastRow, int firstCol, int lastCol, SpreadsheetCellStyle style) {
		for (int row = firstRow; row <= lastRow; row++) {
			for (int col = firstCol; col <= lastCol; col++) {
				getCell(row, col).setStyle(style);
			}
		}
	}

	public SpreadsheetCell getCell(String cellAddress) {
		CellReference cellReference = new CellReference(cellAddress);
		return getCell(cellReference.getRow(), cellReference.getCol());
	}

	public SpreadsheetCell getCell(int row, int col) {
		String address = getCellAddress(row, col);
		SpreadsheetCell cell = cells.get(address);
		if (cell == null) {
			cell = new SpreadsheetCell(this, getPoiCell(row, col));
			cells.put(address, cell);
		}
		return cell;
	}

	public void mergeCells(String firstCell, String lastCell, Object content, SpreadsheetCellStyle style) {
		CellReference firstReference = new CellReference(firstCell);
		CellReference lastReference = new CellReference(lastCell);
		mergeCells(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), content, style);
	}

	public void mergeCells(int firstRow, int lastRow, int firstCol, int lastCol, Object content, SpreadsheetCellStyle style) {
		setValue(firstRow, firstCol, content);
		for (int col = firstCol; col <= lastCol; col++) {
			for (int row = firstRow; row <= lastRow; row++) {
				setStyle(row, col, style);
			}
		}
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * In twips (1/20th of a point)
	 */
	public int getRowHeight(int row) {
		return sheet.getRow(row).getHeight();
	}

	/**
	 * In twips (1/20th of a point)
	 */
	public void setRowHeight(int row, int twips) {
		sheet.getRow(row).setHeight((short)twips);
	}

	/**
	 * In (1/256th of a character width)
	 */
	public int getColWidth(int col) {
		return sheet.getColumnWidth(col);
	}

	/**
	 * In (1/256th of a character width)
	 */
	public void setColWidth(int col, int twips) {
		sheet.setColumnWidth(col, twips);
	}

	public void autosizeRowsAndCols() {
		for (int col = 0; col <= highestModifiedCol; col++) {
			sheet.autoSizeColumn(col, true);
		}
		for (int row = 0; row <= highestModifiedRow; row++) {
			autoSizeRow(row);
		}
	}

	public void autosizeRows() {
		for (int row = 0; row <= highestModifiedRow; row++) {
			autoSizeRow(row);
		}
	}
	
	public void autosizeCols() {
		for (int col = 0; col <= highestModifiedCol; col++) {
			sheet.autoSizeColumn(col, true);
		}
	}

	public void autoSizeRow(int row) {
		float tallestCell = -1;
		for (int col = 0; col <= highestModifiedCol; col++) {
			SpreadsheetCell cell = getCell(row, col);
			int fontSize = cell.getFontSizeInPoints();
			Cell poiCell = cell.getPoiCell();
			if (poiCell.getCellType() == Cell.CELL_TYPE_STRING) {
				String value = poiCell.getStringCellValue();
				int numLines = 1;
				for (int i = 0; i < value.length(); i++) {
					if (value.charAt(i) == '\n') numLines++;
				}
				float cellHeight = computeRowHeightInPoints(fontSize, numLines);
				if (cellHeight > tallestCell) {
					tallestCell = cellHeight;
				}
			}
		}

		float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
		float rowHeight = tallestCell;
		if (rowHeight < defaultRowHeightInPoints+1) {
			rowHeight = -1;	// resets to the default
		}

		sheet.getRow(row).setHeightInPoints(rowHeight);
	}

	public float computeRowHeightInPoints(int fontSizeInPoints, int numLines) {
		// a crude approximation of what excel does
		float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
		float lineHeightInPoints = 1.3f * fontSizeInPoints;
		if (lineHeightInPoints < defaultRowHeightInPoints + 1) {
			lineHeightInPoints = defaultRowHeightInPoints;
		}
		float rowHeightInPoints = lineHeightInPoints * numLines;
		rowHeightInPoints = Math.round(rowHeightInPoints * 4) / 4f;		// round to the nearest 0.25
		return rowHeightInPoints;
	}

	public void addSpacer() {
		sheet.setColumnWidth(0, 768);
	}

	@SuppressWarnings("UnusedDeclaration")
	public Sheet getPoiSheet() {
		return sheet;
	}

	public Cell getPoiCell(int rowNum, int col) {
		Row row = getPoiRow(rowNum);
		Cell cell = row.getCell(col);
		if (cell == null) {
			cell = row.createCell(col);
		}
		return cell;
	}

	private Row getPoiRow(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		return row;
	}

	private void recordCellModified(int row, int col) {
		if (col > highestModifiedCol) {
			highestModifiedCol = col;
		}
		if (row > highestModifiedRow) {
			highestModifiedRow = row;
		}
	}

	public void printDown(String cellAddress, SpreadsheetCellStyle style, String... values) {
		CellReference cellReference = new CellReference(cellAddress);
		printDown(cellReference.getRow(), cellReference.getCol(), style, values);
	}

	public void printAcross(String cellAddress, SpreadsheetCellStyle style, String... values) {
		CellReference cellReference = new CellReference(cellAddress);
		printAcross(cellReference.getRow(), cellReference.getCol(), style, values);
	}

	/**
	 * Returns the index of the next row after the last one written
	 */
	public int printDown(int row, int col, SpreadsheetCellStyle style, Object... values) {
		for (int i = 0; i < values.length; i++) {
			setValue(row + i, col, values[i], style);
		}
		return row + values.length;
	}

	/**
	 * Returns the index of the next col after the last one written.
	 */
	public int printAcross(int row, int col, SpreadsheetCellStyle style, Object... values) {
		for (int i = 0; i < values.length; i++) {
			setValue(row, col + i, values[i], style);
		}
		return col + values.length;
	}

	public void setSurroundBorder(String firstCell, String lastCell, short border) {
		CellReference firstReference = new CellReference(firstCell);
		CellReference lastReference = new CellReference(lastCell);
		setSurroundBorder(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), border);
	}

	public void setSurroundBorder(int firstRow, int lastRow, int firstCol, int lastCol, short border) {
		setTopBorder(firstRow, firstCol, lastCol, border);
		setBottomBorder(lastRow, firstCol, lastCol, border);
		setLeftBorder(firstRow, lastRow, firstCol, border);
		setRightBorder(firstRow, lastRow, lastCol, border);
	}

	public void setRightBorder(int firstRow, int lastRow, int col, short border) {
		for (int row = firstRow; row <= lastRow; row++) {
			getCell(row, col).applyStyle(DEFAULT.withRightBorder(border));
		}
	}

	public void setLeftBorder(int firstRow, int lastRow, int col, short border) {
		for (int row = firstRow; row <= lastRow; row++) {
			getCell(row, col).applyStyle(DEFAULT.withLeftBorder(border));
		}
	}

	public void setTopBorder(int row, int firstCol, int lastCol, short border) {
		for (int col = firstCol; col <= lastCol; col++) {
			getCell(row, col).applyStyle(DEFAULT.withTopBorder(border));
		}
	}

	public void setBottomBorder(int row, int firstCol, int lastCol, short border) {
		for (int col = firstCol; col <= lastCol; col++) {
			getCell(row, col).applyStyle(DEFAULT.withBottomBorder(border));
		}
	}

	public CellStyle registerStyle(SpreadsheetCellStyle style) {
		return workbook.registerStyle(style);
	}

	public void addPicture(String cellAddress, byte[] bytes, int pictureType) {
		CellReference cellRef = new CellReference(cellAddress);
		addPicture(cellRef.getRow(), cellRef.getCol(), bytes, pictureType);
	}

	public void addPicture(int row, int col, byte[] bytes, int pictureType) {
		if (drawing == null) {
			drawing = sheet.createDrawingPatriarch();
		}

		int pictureIndex = workbook.getPoiWorkbook().addPicture(bytes, pictureType);
		//add a picture shape
		ClientAnchor anchor = workbook.getPoiWorkbook().getCreationHelper().createClientAnchor();
		//set top-left corner of the picture,
		//subsequent call of Picture#resize() will operate relative to it
		anchor.setCol1(col);
		anchor.setRow1(row);

		Picture pict = drawing.createPicture(anchor, pictureIndex);
		//auto-size picture relative to its top-left corner
		pict.resize();
	}

}
