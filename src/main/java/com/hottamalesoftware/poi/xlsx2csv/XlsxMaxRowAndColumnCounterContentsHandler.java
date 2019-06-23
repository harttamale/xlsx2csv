package com.hottamalesoftware.poi.xlsx2csv;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Uses the XSSF Event SAX helpers to do most of the work
 * of parsing the Sheet XML, and counts the max column count.
 */
public class XlsxMaxRowAndColumnCounterContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private static final char END_OF_LINE = '\n';
    private boolean firstCellOfRow = false;
    private int currentRow = -1;
    private int currentCol = -1;
    private int maxRow;
    private int maxColumn;

    public XlsxMaxRowAndColumnCounterContentsHandler() {
    }

    public int getMaxColumn() {
        return maxColumn + 1;//for offset
    }

    public int getMaxRow() {
        return maxRow + 1;//for offset
    }

    @Override
    public void startRow(int rowNum) {
        currentRow = rowNum;
        currentCol = -1;
        if (rowNum > maxRow) {
            maxRow = rowNum;
        }
    }

    @Override
    public void endRow(int rowNum) {
        if (currentCol > maxColumn) {
            maxColumn = currentCol;
        }
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        // gracefully handle missing CellRef here in a similar way as XSSFCell does
        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        // Did we miss any cells?
        int thisCol = (new CellReference(cellReference)).getCol();
        currentCol = thisCol;
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // Skip, no headers or footers in CSV
    }

}