# fromExcelToCSV

A terminal app written in Java that reads Excel rows and export to a pipe-delimited csv file.

Developed using Visual Studio Code and Apache Maven with the following extensions:
- Maven for Java
- Language Support for Java(TM) by Red Hat
- Debugger for Java
- Java Test Runner
- Java Extension Pack
- Java Dependency Viewer

## Source file requirements
1. The app only accepts Excel 2007 (or later) files in *.xlsx
2. Source must be placed on the same directory as the java (or jar) file
3. To modify default filename, update `INPUT_FILENAME` variable
4. First sheet name must be a month name (e.g. January)
5. Year must be placed at `E3` cell (e.g. 2018)
6. Type of menu must be placed at `A` cell on any row (e.g. International)
7. Data rows are read from `A` to `E` cells where `A` cell values must be numeric

## How to run
```
$ java -cp fromExcelToCSV.jar org.sample.fretlesskey.App
```

## Output: csv and log files
- YYYYMMSRC_FILENAME.csv
(e.g. 201804FOOD_MENU.csv)
- YYYYMMDD_HHMMHH_fromExcelToCSV.py.log
(e.g. 20180402_080012_fromExcelToCSV.py.log)