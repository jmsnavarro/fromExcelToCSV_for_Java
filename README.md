# fromExcelToCSV in Java

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

> Note: See **Developer Notes** below to create the jar file

```
## simple
$ java -cp fromExcelToCSV.jar org.sample.fretlesskey.App
```
```
## with output log (in Powershell terminal)
$ java -cp fromExcelToCSV.jar org.sample.fretlesskey.App > "$(Get-Date -f yyyMMdd_HHmm)_fromExcelToCSV.java.log"
```
```
## with output log (in Linux terminal)
$ java -cp fromExcelToCSV.jar org.sample.fretlesskey.App > "$(date +%Y%M%d_%H%M)_fromExcelToCSV.java.log"
```

## Output

**CSV**

Format: YYYYMMSRC_FILENAME.csv
```
201804FOOD_MENU.csv
```
**Log**

Format: YYYYMMDD_HHMMHH_fromExcelToCSV.java.log
```
20190712_1402_fromExcelToCSV.java.log
```

## Developer Notes:

```
## generate single jar file
$ mvn clean compile assembly:single
```
```
## cleaning up
$ mvn clean
```