# fromExcelToCSV in Java

A command-line app written in Java that reads Excel rows and export to a pipe-delimited csv file.

Developed using Visual Studio Code and Apache Maven with the following extensions:

- Maven for Java
- Language Support for Java(TM) by Red Hat
- Debugger for Java
- Java Test Runner
- Java Extension Pack
- Java Dependency Viewer

Tested environment

- Windows Subsystem for Linux (WSL) v1 using Ubuntu 18.04 LTS
- openjdk v11.0.6
- Visual Studio Code v1.41.1 x64

## Source file requirements

1. The app **only** accepts Excel 2007 (or later) files in *.xlsx, **not** *.xls (Excel 2000, 97 or earlier)
2. Source file can be placed on the same directory as the java (or jar) file or in another path
3. First sheet name must be a month name (e.g. `January`)
4. Year must be placed at `E3` cell (e.g. 2018)
5. Type of menu must be placed at `A` cell on any row (e.g. `International`)
6. Data rows are read from `A` to `E` cells where `A` cell values must be numeric

## How to run

> Note: See **Developer Notes** below to create the jar file

```bash
## simple (log file generation included)
$ java -jar fromExcelToCSV.jar ~/src/myfile.xlsx off
```

```bash
## with additional output log (in Powershell terminal)
$ java -jar fromExcelToCSV.jar ~/src/myfile.xlsx off > "$(Get-Date -f yyyMMdd_HHmm)_fromExcelToCSV.java.log"
```

```bash
## with additional output log (in Linux terminal)
$ java -jar fromExcelToCSV.jar ~/src/myfile.xlsx off > "$(date +%Y%m%d_%H%M%S)_fromExcelToCSV.java.log"
```

## Output

### CSV

Format: YYYYMMSRC_FILENAME.csv

```bash
201804FOOD_MENU.csv
```

### Log

Format: YYYYMMDD_HHMMHH_fromExcelToCSV.java.log

```bash
20190712_1402_fromExcelToCSV.java.log
```

## Developer Notes

```bash
## To generate single jar file
$ mvn clean compile assembly:single
```

```bash
## To clean up
$ mvn clean
```
