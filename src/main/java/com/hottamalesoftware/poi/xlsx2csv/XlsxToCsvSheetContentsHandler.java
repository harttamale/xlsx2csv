package com.hottamalesoftware.poi.xlsx2csv;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.io.PrintStream;

/**
 * Uses the XSSF Event SAX helpers to do most of the work
 * of parsing the Sheet XML, and outputs the contents
 * as a (basic) CSV.
 */
public class XlsxToCsvSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private static final char END_OF_LINE = '\n';
    private boolean firstCellOfRow = false;
    private int currentRow = -1;
    private int currentCol = -1;
    private XlsxConverterState xlsxConverterState;

    public XlsxToCsvSheetContentsHandler(XlsxConverterState xlsxConverterState) {
        this.xlsxConverterState = xlsxConverterState;
    }


    @Override
    public void startRow(int rowNum) {
        // If there were gaps, output the missing rows
        outputMissingRows(rowNum - currentRow - 1);
        // Prepare for this row
        firstCellOfRow = true;
        currentRow = rowNum;
        currentCol = -1;
    }

    @Override
    public void endRow(int rowNum) {
        PrintStream output = xlsxConverterState.getOutput();
        // Ensure the minimum number of columns
        for (int i = currentCol; i < xlsxConverterState.getMinColumns(); i++) {
            output.append(',');
        }
        output.append(END_OF_LINE);
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        PrintStream output = xlsxConverterState.getOutput();
        if (firstCellOfRow) {
            firstCellOfRow = false;
        } else {
            output.append(',');
        }

        // gracefully handle missing CellRef here in a similar way as XSSFCell does
        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        // Did we miss any cells?
        int thisCol = (new CellReference(cellReference)).getCol();
        int missedCols = thisCol - currentCol - 1;
        for (int i = 0; i < missedCols; i++) {
            output.append(',');
        }
        currentCol = thisCol;

        // Number or string?
        try {
            Double.parseDouble(formattedValue);
            output.append(formattedValue);
        } catch (NumberFormatException e) {
            output.append('"');
            output.append(formattedValue);
            output.append('"');
        }
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // Skip, no headers or footers in CSV
    }

    private void outputMissingRows(int number) {
        PrintStream output = xlsxConverterState.getOutput();
        for (int i = 0; i < number; i++) {
            for (int j = 0; j < xlsxConverterState.getMinColumns(); j++) {
                output.append(',');
            }
            output.append(END_OF_LINE);
        }
    }
}