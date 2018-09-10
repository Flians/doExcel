
# The implements of parsing Excel(.xlsx and .xls)
  obtain the information of .xlsx File by using streaming method. <br>
  
  It's based on org.apache.poi, override the methods of poi. Therefore you could use this like poi. <br>
  It's a good solution for the memory overflow problem. <br>
  
  The example of using this: <br>
  ```java
  File file = new File("/path/to/workbook.xlsx");
  Workbook workbook = StreamingReader.builder()             // new a Builder object to parse Excel
                                     .sstCacheSize(-1)      // number of rows to keep in memory for the SharedString table (defaults to 10, -1 represents keeping all in memory)
                                     .rowCacheSize(10)      // number of rows to keep in memory (defaults to 10, greater than 0)
                                     .sheetIndex(-1)        // index of sheet to use (defaults to -1, representing that you can read all sheets)
                                     .open(file);           // File for XLSX/XLS file (required)
  ```                             
  Or (Not recommended)
  ```java
  InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
  Workbook workbook = StreamingReader.builder()             // new a Builder object to parse Excel
                                     .sstCacheSize(-1)      // number of rows to keep in memory for the SharedString table (defaults to 10, -1 represents keeping all in memory)
                                     .rowCacheSize(10)      // number of rows to keep in memory (defaults to 10)
									 .bufferSize(1024)      // buffer size to use when reading InputStream to file (defaults to 1024)
                                     .sheetIndex(-1)        // index of sheet to use (defaults to -1, representing that you can read all sheets)
                                     .open(is, excelType);  // InputStream for XLSX/XLS file (required)
  ```