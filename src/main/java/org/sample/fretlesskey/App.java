package org.sample.fretlesskey;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.commons.validator.GenericValidator;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A terminal app written in Java that reads Excel rows and exports to a pipe-delimited csv file.
 *
 * Developed using Visual Studio Code and Apache Maven with the following extensions:
 * - Maven for Java
 * - Language Support for Java(TM) by Red Hat
 * - Debugger for Java
 * - Java Test Runner
 * - Java Extension Pack
 * - Java Dependency Viewer
 *
 */

//TODO: Write execution steps to a log file
//TODO: Add Excel file version detection (2003 vs 2007) using Apache Tika

public class App
{
    //constants
    static final String PLATTER = "PLATTER";
    static final String DRINKS = "DRINKS";
    static final String INPUT_FILENAME = "food_menu.xlsx";
    static final String YEAR_COLUMN = "E3";
    static final String MESSAGE_NO_VALID_YEAR_ON_HEADER =
        "No defined year value in cell 'E3'. Kindly check the source file.";
    static final String OUTPUT_FILE = "FOOD_MENU";

    //Logging
    private static final Logger logger = LogManager.getLogger(App.class);

    //Year validator
    public static boolean yearIsValid(String year) {
        return GenericValidator.isDate(year, "yyyy", true);
    }

    //Get cell value
    public static String getCellValue(Cell cell) {
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
    public static String getNumMonthVal(String monthName) {
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

    public static void main( String[] args ) {
        //Variables
        String strCellValue = null;
        Integer intYear = 0;
        String strSheetName = null;
        String monthNumber = null;
        String strMenuGroup = null;

        try {
            //Get Excel file
            FileInputStream file = new FileInputStream(new File(INPUT_FILENAME));
            logger.info("Opening file '" + INPUT_FILENAME + "'");

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Get sheet name
            strSheetName = sheet.getSheetName().toString();
            monthNumber = getNumMonthVal(strSheetName);

            //Build CSV
            try {
                logger.info("Preparing to read your file");

                //Get year else app exit
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
                        logger.warn(e.getMessage());
                    }
                } else {
                    logger.warn(MESSAGE_NO_VALID_YEAR_ON_HEADER);
                    workbook.close();
                    file.close();
                    System.exit(0);
                }

                //Generate output filename
                StringBuilder sbOutputFilename = new StringBuilder();
                sbOutputFilename.append(intYear.toString());
                sbOutputFilename.append(monthNumber);
                sbOutputFilename.append(OUTPUT_FILE);
                sbOutputFilename.append(".csv");

                FileWriter csvFileWriter =
                    new FileWriter(sbOutputFilename.toString());
                CSVFormat csvFileFormat = CSVFormat.DEFAULT.withDelimiter('|');
                CSVPrinter csvFilePrinter = new CSVPrinter(csvFileWriter, csvFileFormat);
                ArrayList<Object> detail = new ArrayList<Object>();

                //Iterate through each rows one by one
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    logger.info("Reading row " + (row.getRowNum() - 1));

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
                    }
                    detail.clear();
                }
                csvFilePrinter.close();

            } catch (IOException e) {
                workbook.close();
                file.close();
                logger.warn(e.getMessage());
            }
        } catch (Exception e) {
            logger.warn(e.getMessage());
            System.exit(0);
        }
        logger.info("Successfully created a CSV file");
        System.exit(0);
    }
}
