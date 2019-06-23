/* ====================================================================
Licensed to the Apache Software Foundation (ASF) under one or more
contributor license agreements.  See the NOTICE file distributed with
this work for additional information regarding copyright ownership.
The ASF licenses this file to You under the Apache License, Version 2.0
(the "License"); you may not use this file except in compliance with
the License.  You may obtain a copy of the License at
   http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
==================================================================== */
package com.hottamalesoftware.poi.xlsx2csv;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;

/**
 * A rudimentary XLSX -> CSV processor modeled on the
 * POI sample program XLS2CSVmra from the package
 * org.apache.poi.hssf.eventusermodel.examples.
 * As with the HSSF version, this tries to spot missing
 * rows and cells, and output empty entries for them.
 * <p/>
 * Data sheets are read using a SAX parser to keep the
 * memory footprint relatively small, so this should be
 * able to read enormous workbooks.  The styles table and
 * the shared-string table must be kept in memory.  The
 * standard POI styles table class is used, but a custom
 * (read-only) class is used for the shared string table
 * because the standard POI SharedStringsTable grows very
 * quickly with the number of unique strings.
 * <p/>
 * For a more advanced implementation of SAX event parsing
 * of XLSX files, see {@link XSSFEventBasedExcelExtractor}
 * and {@link XSSFSheetXMLHandler}. Note that for many cases,
 * it may be possible to simply use those with a custom
 * {@link SheetContentsHandler} and no SAX code needed of
 * your own!
 */
public class XlsxToCsvConverter {

    private final XlsxConverterState xlsxConverterState;

    /**
     * Creates a new XLSX -> CSV converter
     *
     * @param input      The input File to process
     * @param output     The PrintStream to output the CSV to
     * @param minColumns The minimum number of columns to output, or -1 for no minimum
     */
    public XlsxToCsvConverter(File input, PrintStream output, int minColumns, boolean autoColumns) {
        this.xlsxConverterState = new XlsxConverterState(input, output, minColumns, autoColumns);
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles
     * @param strings
     * @param sheetInputStream
     */
    public void processSheet(
            StylesTable styles,
            ReadOnlySharedStringsTable strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    public void process()
            throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        doCountColumns();
        doProcessFile();
    }

    private void doCountColumns() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        if (xlsxConverterState.isAutoColumns()) {
            OPCPackage opcPackage = OPCPackage.open(xlsxConverterState.getInputFile(), PackageAccess.READ);
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opcPackage);
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int index = 0;
            if (iter.hasNext()) {//found sheet
                InputStream stream = iter.next();
                String sheetName = iter.getSheetName();
                XlsxMaxRowAndColumnCounterContentsHandler maxRowAndColumnHandler = new XlsxMaxRowAndColumnCounterContentsHandler();
                processSheet(styles, strings, maxRowAndColumnHandler, stream);
                this.xlsxConverterState.getOutput().println("xlsx2csv parser max rows " + maxRowAndColumnHandler.getMaxRow() + " max columns " + maxRowAndColumnHandler.getMaxColumn());
                this.xlsxConverterState.setAutoColumnWidth(maxRowAndColumnHandler.getMaxColumn());
                stream.close();
                ++index;
            }
            opcPackage.close();
        }
    }

    private void doProcessFile() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        OPCPackage opcPackage = OPCPackage.open(xlsxConverterState.getInputFile(), PackageAccess.READ);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opcPackage);
        XSSFReader xssfReader = new XSSFReader(opcPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        if (iter.hasNext()) {//found sheet
            InputStream stream = iter.next();
            String sheetName = iter.getSheetName();
            this.xlsxConverterState.getOutput().println("xlsx2csv parser result");
            this.xlsxConverterState.getOutput().println(sheetName + " [index=" + index + "]:");
            processSheet(styles, strings, new XlsxToCsvSheetContentsHandler(xlsxConverterState), stream);
            stream.close();
            ++index;
        }
        opcPackage.close();
    }
}
