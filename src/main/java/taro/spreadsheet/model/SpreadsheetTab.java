package taro.spreadsheet.model;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Map;

import static com.google.common.collect.Maps.newHashMap;
import static taro.spreadsheet.model.SpreadsheetCellStyle.DEFAULT;

public class SpreadsheetTab {

    private SpreadsheetWorkbook workbook;
    private XSSFSheet sheet;
    private Map<String, SpreadsheetCell> cells = newHashMap();
    private Drawing drawing;

    private int highestModifiedCol = -1;
    private int highestModifiedRow = -1;

    SpreadsheetTab(SpreadsheetWorkbook workbook, String title) {
        this.workbook = workbook;
        this.sheet = workbook.getPoiWorkbook().createSheet(title);
    }

    SpreadsheetTab(SpreadsheetWorkbook workbook, XSSFSheet sheet) {
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
        SpreadsheetCell cell = getOrCreateCell(row, col);
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
        getOrCreateCell(row, col).setStyle(style);
    }

    public void setStyle(int firstRow, int lastRow, int firstCol, int lastCol, SpreadsheetCellStyle style) {
        for (int row = firstRow; row <= lastRow; row++) {
            for (int col = firstCol; col <= lastCol; col++) {
                getOrCreateCell(row, col).setStyle(style);
            }
        }
    }

    public XSSFSheet getPoiSheet() {
        return sheet;
    }

    public SpreadsheetCell getCell(String cellAddress) {
        CellReference cellReference = new CellReference(cellAddress);
        return getCell(cellReference.getRow(), cellReference.getCol());
    }

    public SpreadsheetCell getCell(int row, int col) {
        String address = getCellAddress(row, col);
        SpreadsheetCell cell = cells.get(address);
        return cell;
    }

    public SpreadsheetCell getOrCreateCell(String cellAddress) {
        CellReference cellReference = new CellReference(cellAddress);
        return getOrCreateCell(cellReference.getRow(), cellReference.getCol());
    }

    public SpreadsheetCell getOrCreateCell(int row, int col) {
        SpreadsheetCell cell = getCell(row,col);
        if (cell == null) {
            cell = new SpreadsheetCell(this, getOrCreatePoiCell(row, col));
            String address = getCellAddress(row, col);
            cells.put(address, cell);
        }
        return cell;
    }

    private XSSFCell getOrCreatePoiCell(int rowNum, int col) {
        XSSFRow row = getOrCreatePoiRow(rowNum);
        XSSFCell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        return cell;
    }

    private XSSFRow getOrCreatePoiRow(int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
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
        autosizeCols();
        autosizeRows();
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
            SpreadsheetCell cell = getOrCreateCell(row, col);
            int fontSize = cell.getFontSizeInPoints();
            XSSFCell poiCell = cell.getPoiCell();
            if (poiCell.getCellType() == CellType.STRING) {
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
            rowHeight = -1;    // resets to the default
        }

        sheet.getRow(row).setHeightInPoints(rowHeight);
    }

    public float computeRowHeightInPoints(int fontSizeInPoints, int numLines) {
        // a crude approximation of what excel does
        float lineHeightInPoints = 1.3f * fontSizeInPoints;
        float rowHeightInPoints = lineHeightInPoints * numLines;
        rowHeightInPoints = Math.round(rowHeightInPoints * 4) / 4f;        // round to the nearest 0.25

        // Don't shrink rows to fit the font, only grow them
        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        if (rowHeightInPoints < defaultRowHeightInPoints + 1) {
            rowHeightInPoints = defaultRowHeightInPoints;
        }
        return rowHeightInPoints;
    }

    public void addSpacer() {
        sheet.setColumnWidth(0, 768);
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

    public void setSurroundBorder(String firstCell, String lastCell, BorderStyle border) {
        CellReference firstReference = new CellReference(firstCell);
        CellReference lastReference = new CellReference(lastCell);
        setSurroundBorder(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), border);
    }

    public void setSurroundBorder(int firstRow, int lastRow, int firstCol, int lastCol, BorderStyle border) {
        setTopBorder(firstRow, firstCol, lastCol, border);
        setBottomBorder(lastRow, firstCol, lastCol, border);
        setLeftBorder(firstRow, lastRow, firstCol, border);
        setRightBorder(firstRow, lastRow, lastCol, border);
    }

    public void setRightBorder(int firstRow, int lastRow, int col, BorderStyle border) {
        for (int row = firstRow; row <= lastRow; row++) {
            getOrCreateCell(row, col).applyStyle(DEFAULT.withRightBorder(border));
        }
    }

    public void setLeftBorder(int firstRow, int lastRow, int col, BorderStyle border) {
        for (int row = firstRow; row <= lastRow; row++) {
            getOrCreateCell(row, col).applyStyle(DEFAULT.withLeftBorder(border));
        }
    }

    public void setTopBorder(int row, int firstCol, int lastCol, BorderStyle border) {
        for (int col = firstCol; col <= lastCol; col++) {
            getOrCreateCell(row, col).applyStyle(DEFAULT.withTopBorder(border));
        }
    }

    public void setBottomBorder(int row, int firstCol, int lastCol, BorderStyle border) {
        for (int col = firstCol; col <= lastCol; col++) {
            getOrCreateCell(row, col).applyStyle(DEFAULT.withBottomBorder(border));
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
