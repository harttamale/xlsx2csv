package com.hottamalesoftware.poi.xlsx2csv;

import org.apache.poi.openxml4j.opc.OPCPackage;

import java.io.File;
import java.io.PrintStream;

public class XlsxConverterState {
    private final File inputFile;
    /**
     * Number of columns to read starting with leftmost
     */
    private final int minColumns;
    /**
     * Destination for data
     */
    private final PrintStream output;

    /**
     * Should max column width be counted
     */
    private final boolean autoColumns;

    private int autoColumnWidth;


    public XlsxConverterState(File inputFile, PrintStream output, int minColumns, boolean autoColumns) {
        this.inputFile=inputFile;
        this.output = output;
        this.minColumns = minColumns;
        this.autoColumns = autoColumns;
    }

    public boolean isAutoColumns() {
        return autoColumns;
    }

    public File getInputFile() {
        return inputFile;
    }

    public int getMinColumns() {
        return autoColumns ? autoColumnWidth : minColumns;
    }

    public PrintStream getOutput() {
        return output;
    }

    public void setAutoColumnWidth(int autoColumnWidth) {
        this.autoColumnWidth = autoColumnWidth;
    }
}
