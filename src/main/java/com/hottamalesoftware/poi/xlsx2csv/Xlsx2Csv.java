package com.hottamalesoftware.poi.xlsx2csv;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.Banner;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.Arrays;
import java.util.List;

@SpringBootApplication
public class Xlsx2Csv implements ApplicationRunner {
    private static final Logger logger = LoggerFactory.getLogger(Xlsx2Csv.class);

    public static void main(String[] args) {

        new SpringApplicationBuilder(Xlsx2Csv.class)
                .bannerMode(Banner.Mode.OFF)
                .logStartupInfo(false)
                .build()
                .run(args);
    }


    @Override
    public void run(ApplicationArguments args) throws Exception {
        long start = System.currentTimeMillis();

        logger.debug("Application started with command-line arguments: {}", Arrays.toString(args.getSourceArgs()));
        logger.debug("NonOptionArgs: {}", args.getNonOptionArgs());

        for (String name : args.getOptionNames()) {
            logger.info("Options: arg-" + name + "=" + args.getOptionValues(name));
        }

        List<String> nonOptionArgs = args.getNonOptionArgs();
        int nonOptionArgsCount = nonOptionArgs.size();
        if (nonOptionArgsCount < 1 || nonOptionArgsCount > 2) {
            System.err.println("Use:");
            System.err.println("  xlsx2csv <input xlsx file> [output csv file]  [--options]");
            System.err.println("Where: ");
            System.err.println("  <xlsx file>                 - Required, this is the file to transform from xlsx to csv");
            System.err.println("  [output csv file]           - Not required, if provided the output file will be created, otherwise will parse to console");
            System.err.println("Options: ");
            System.err.println("  --min-columns <minColWidth> - Not required, ensure min column width when parsing, default -1 which will dump csv rows with no minimum column width");
            System.err.println("  --auto-columns              - Not required, not yet implemented");
            System.err.println("  --sheet-number              - Not required, not yet implemented");
            System.err.println("  --sheet-name                - Not required, not yet implemented");
            System.err.println("  --all-sheets                - Not required, not yet implemented");
            return;
        }

        String inputXlsxFilename = nonOptionArgs.get(0);
        File inputFile = new File(inputXlsxFilename);
        if (!inputFile.exists()) {
            System.err.println("Not found or not a file: " + inputFile.getPath());
            return;
        }

        PrintStream outputStream = null;
        if (nonOptionArgsCount > 1) {
            outputStream = new PrintStream(new FileOutputStream(inputFile.getName().substring(0, inputFile.getName().lastIndexOf('.')), true));
        } else {
            outputStream = System.out;
        }

        String minColumnsValue = getOne(args, "min-columns");
        int minColumns = -1;
        if (minColumnsValue != null) {
            minColumns = Integer.parseInt(minColumnsValue);
        }

        boolean autoColumns = isSet(args, "auto-columns");

        XlsxToCsvConverter xlsx2csv = new XlsxToCsvConverter(inputFile, outputStream, minColumns, autoColumns);
        xlsx2csv.process();

        System.out.println("Completed XLSX2CSV in " + (System.currentTimeMillis() - start) + " ms.");
    }

    private boolean isSet(ApplicationArguments applicationArguments, String optionName) {
        return applicationArguments.getOptionNames().contains(optionName);
    }

    // Has exactly on option value
    private String getOne(ApplicationArguments applicationArguments, String optionName) {
        List<String> values = applicationArguments.getOptionValues(optionName);
        if (values == null) {
            return null;
        }
        if (values.size() > 1 || values.size() == 0) {
            System.err.println("Invalid option --" + optionName + ", ignoring");
            return null;
        } else {
            return values.get(0);
        }

    }
}
