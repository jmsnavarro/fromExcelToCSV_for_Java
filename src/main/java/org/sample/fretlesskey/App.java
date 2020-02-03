package org.sample.fretlesskey;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.time.Duration;
import java.time.LocalDateTime;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.lookup.*;
import org.apache.commons.validator.GenericValidator;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.text.CaseUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.Tika;

/*
 * A terminal app written in Java that reads Excel rows and exports to a pipe-delimited csv file.
 *
 * Developed using Visual Studio Code and Apache Maven with the following extensions:
 * - Maven for Java
 * - Language Support for Java(TM) by Red Hat
 * - Debugger for Java
 * - Java Test Runner
 * - Java Extension Pack
 * - Java Dependency Viewer
 */

public class App
{
    //constants
    static final String APPNAME = App.class.getName();
    static final String MIMETYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static final String PLATTER = "PLATTER";
    static final String DRINKS = "DRINKS";
    static final String YEAR_COLUMN = "E3";
    static final String MESSAGE_APP_WILL_NOW_CLOSE = "App will now close.";
    static final String MESSAGE_NO_VALID_YEAR_ON_HEADER =
        "No defined year value in cell 'E3'. Kindly check the source file.";
    static final String OUTPUT_FILE = "FOOD_MENU";

    //Logging
    //TODO: Add parameter-based enable/disable log4j2 RollingFile appender 
    //      https://logging.apache.org/log4j/2.x/manual/lookups.html
    private static final Logger logger = LogManager.getLogger(App.class);

    //Year validator
    private static boolean yearIsValid(String year) {
        return GenericValidator.isDate(year, "yyyy", true);
    }

    //Get cell value
    private static String getCellValue(Cell cell) {
        String strCellValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                strCellValue = Double.toString(cell.getNumericCellValue());
                break;
            case STRING:
                strCellValue = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return strCellValue;
    }

    //Format month name to month number
    private static String getNumMonthVal(String monthName) {
        String month = monthName.toUpperCase();
        if (month.equals("JAN") || month.equals("JANUARY")) {
            return "01";
        } else if (month.equals("FEB") || month.equals("FEBRUARY")) {
            return "02";
        } else if (month .equals("MAR") || month.equals("MARCH")) {
            return "03";
        } else if (month.equals("APR") || month.equals("APRIL")) {
            return "04";
        } else if (month.equals("MAY")) {
            return "05";
        } else if (month.equals("JUN") || month.equals("JUNE")) {
            return "06";
        } else if (month.equals("JUL") || month.equals("JULY")) {
            return "07";
        } else if (month.equals("AUG") || month.equals("AUGUST")) {
            return "08";
        } else if (month.equals("SEP") || month.equals("SEPTEMBER")) {
            return "09";
        } else if (month.equals("OCT") || month.equals("OCTOBER")) {
            return "10";
        } else if (month.equals("NOV") || month.equals("NOVEMBER")) {
            return "11";
        } else if (month.equals("DEC") || month.equals("DECEMBER")) {
            return "12";
        } else {
            return "00";
        }
    }

    //Detect document (stream) type
    private static String getMimeType(InputStream stream) throws IOException {
        Tika tika = new Tika();
        String mimeType = tika.detect(stream);
        return mimeType;
    }

    //Check if file exists
    private static boolean ifFileExists(String file) {
        boolean exists = false;
        if (file.trim().length() > 0) {
            try {
                exists = Paths.get(file).toFile().exists() ? true : false;
            } catch (Exception e) {
                logger.error(e.getMessage());
                logger.info(MESSAGE_APP_WILL_NOW_CLOSE);
                System.exit(0);
            }
        } else {
            exists = false;
        }
        return exists;
    }

    //Validate Excel (.xlsx) file
    private static boolean isExcelFile(String file) {
        boolean isExcel = false;
        try {
            logger.info("Validating '" + file + "' source file...");
            FileInputStream inputFileInputStream = new FileInputStream(file);
            String mimeType = getMimeType(inputFileInputStream);     
            isExcel = (mimeType.equals(MIMETYPE_XLSX)) ? true : false; 
        } catch (Exception e) {
            logger.error(e.getMessage());
            System.exit(0);
        }
        return isExcel;
    }

    public static void main( String[] args ) {
        //Variables
        String strFileName = "food_menu.xlsx";
        String strCellValue = null;
        Integer intYear = 0;
        String strSheetName = null;
        String strOutputFileName = null;
        String monthNumber = null;
        String strMenuGroup = null;
        Integer intRecordCount = 0;

        //Set run time
        LocalDateTime startDateTime = LocalDateTime.now();

        //Enable/Disable logger (not yet fully implemented)
        MainMapLookup.setMainArguments(args);

        try {
            //Show execution start
            logger.info("Starting execution of " + APPNAME);

            //Check if filename provided by parameter else
            //  will check on the same directory as the app
            if (args[0].toString().trim().length() > 0) {
                strFileName = args[0].toString().trim();
            }

            //Check file if exists
            if (!ifFileExists(strFileName)) {
                logger.error("ERROR: File '" + strFileName + "' do not exists. Please verify.");
                logger.info(MESSAGE_APP_WILL_NOW_CLOSE);
                System.exit(0);
            }

            //Validate if an Excel (.xlsx) file
            if (isExcelFile(strFileName) == true) {
                logger.info("VALIDATED: Source file");
            } else {
                logger.error("'" + strFileName + "' is not a valid Excel (.xlsx) file.");
                logger.info(MESSAGE_APP_WILL_NOW_CLOSE);
                System.exit(0);
            }

            //Get Excel file
            FileInputStream file = new FileInputStream(strFileName);
            
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Get sheet name
            strSheetName = sheet.getSheetName().toString();
            monthNumber = getNumMonthVal(strSheetName);

            //Build CSV
            try {
                //Validate year else app exit
                CellReference cellReference = new CellReference(YEAR_COLUMN);
                Row headerRow = sheet.getRow(cellReference.getRow());
                Cell headerCell = headerRow.getCell(cellReference.getCol());

                if (headerCell.getCellType() == CellType.NUMERIC) {
                    try {
                        intYear = (int)headerCell.getNumericCellValue();
                        if (yearIsValid(intYear.toString()) == false) {
                            logger.warn(MESSAGE_NO_VALID_YEAR_ON_HEADER);
                            workbook.close();
                            file.close();
                            System.exit(0);
                        }
                    } catch (Exception e) {
                        logger.error(e.getMessage());
                    }
                } else {
                    logger.warn(MESSAGE_NO_VALID_YEAR_ON_HEADER);
                    workbook.close();
                    file.close();
                    System.exit(0);
                }

                //Show log header
                logger.info("Opening file '" + strFileName + "' for "
                    + CaseUtils.toCamelCase(strSheetName, true) + " "
                    + intYear + "...");

                //Generate output filename
                StringBuilder sbOutputFilename = new StringBuilder();
                sbOutputFilename.append(intYear.toString());
                sbOutputFilename.append(monthNumber);
                sbOutputFilename.append(OUTPUT_FILE);
                sbOutputFilename.append(".csv");
                strOutputFileName = sbOutputFilename.toString();

                FileWriter csvFileWriter =
                    new FileWriter(strOutputFileName);
                CSVFormat csvFileFormat = CSVFormat.DEFAULT.withDelimiter('|');
                CSVPrinter csvFilePrinter = new CSVPrinter(csvFileWriter, csvFileFormat);
                ArrayList<Object> detail = new ArrayList<Object>();

                //Iterate through each rows one by one
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();

                    //For each row, iterate through all the columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        //Get cell value by menu group
                        strCellValue = getCellValue(cell);
                        if (strCellValue != null) {
                            if (strCellValue.equals(PLATTER)) {
                                strMenuGroup = PLATTER;
                            } else if (strCellValue.equals(DRINKS)) {
                                strMenuGroup = DRINKS;
                            //TODO: Exclude footer. Workaround used.
                            } else if (strCellValue.contains("*")) {
                            //} else {
                                strMenuGroup = null;
                            }
                            detail.add(strCellValue);
                        }
                    }
                    if (strCellValue != null && strMenuGroup != null) {
                        detail.add(0, strSheetName);
                        detail.add(1, intYear);
                        detail.add(2, strMenuGroup);
                        csvFilePrinter.printRecord(detail);

                        intRecordCount++;
                        logger.info("Collecting data from '" + strMenuGroup
                            + "' at row " + (row.getRowNum()));
                    }
                    detail.clear();
                }
                csvFilePrinter.close();

            } catch (IOException e) {
                workbook.close();
                file.close();
                logger.error(e.getMessage());
            }
        } catch (Exception e) {
            logger.error(e.getMessage());
            System.exit(0);
        }

        //Summary
        String formattedRunTime = null;
        try {
            LocalDateTime endDateTime = LocalDateTime.now();
            Duration spanDateTime = Duration.between(startDateTime, endDateTime);
            formattedRunTime = String.format("%d:%02d:%02d",
                spanDateTime.toHours(),
                spanDateTime.toMinutes(),
                spanDateTime.toMillis());
        } catch (Exception e) {
            logger.error(e.getMessage());
        }

        logger.info("SUCCESS: Done copying " + (intRecordCount) + " row(s) to "
            + strOutputFileName
            + " | Elapsed time "
            + formattedRunTime);

        //Close app
        System.exit(0);
    }
}
