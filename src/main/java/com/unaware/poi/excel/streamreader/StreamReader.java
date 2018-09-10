package com.unaware.poi.excel.streamreader;

import com.unaware.poi.excel.exception.NotSupportedException;
import com.unaware.poi.excel.ssimpl.StreamWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Unaware
 * @date 2018/7/12 15:24
 * obtain the information of .xlsx File by using streaming method.
 * It's based on org.apache.poi, override the methods of poi. Therefore you could use this like poi.
 * It's a good solution for the memory overflow problem.
 * The example of using this:
 * <pre>
 * File file = new File("/path/to/workbook.xlsx");
 * Workbook workbook = StreamingReader.builder()
 * .sstCacheSize(-1)      // number of rows to keep in memory for the SharedString table (defaults to 10, -1 represents keeping all in memory)
 * .rowCacheSize(10)      // number of rows to keep in memory (defaults to 10, greater than 0)
 * .sheetIndex(-1)        // index of sheet to use (defaults to -1, representing that you can read all sheets)
 * .open(file);           // File for XLSX/XLS file (required)
 *
 * Or (Not recommended)
 * InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
 * Workbook workbook = StreamingReader.builder()
 * .sstCacheSize(-1)      // number of rows to keep in memory for the SharedString table (defaults to 10, -1 represents keeping all in memory)
 * .rowCacheSize(10)      // number of rows to keep in memory (defaults to 10)
 * .bufferSize(1024)      // buffer size to use when reading InputStream to file (defaults to 1024)
 * .sheetIndex(-1)        // index of sheet to use (defaults to -1, representing that you can read all sheets)
 * .open(is, excelType);  // InputStream for XLSX/XLS file (required)
 * </pre>
 */
public class StreamReader {

    public static Builder builder() {
        return new Builder();
    }

    /**
     * streaming .xlsx File
     */
    public static class Builder {
        public enum ExcelType {XLSX, XLS}

        /**
         * The number of rows to keep in memory
         */
        private int rowCacheSize = 10;
        /**
         * The number of bytes to read once into memory from the inputStream
         */
        private int bufferSize = 1024;
        /**
         * The size of the SharedString table cache.
         * -1 represents that you can keep all in memory
         */
        private int sstCacheSize = 10;
        /**
         * The index of the sheet opened
         * There can only be one sheet open for a single instance of {@link StreamReader}.
         * If you want to read more sheets at the same time, you can create new instances.
         * -1 represents that you can read all sheets
         */
        private int sheetIndex = -1;

        /**
         * The password to unlock this .xlsx file
         */
        private String password;

        /**
         * @param is        文件流
         * @param excelType 文件类型
         * @return org.apache.poi.ss.usermodel.Workbook
         * read a giver {@link InputStream} and return a new instance of {@link Workbook}.
         * a temporary file will be written to create a streaming iterator.
         * In this way, we can read little by little from InputStream.
         * Not recommended
         */
        public Workbook open(InputStream is, ExcelType excelType) throws IOException {
            if (excelType == ExcelType.XLSX) {
                WorkbookReader workbookReader = new WorkbookReader(this);
                workbookReader.init(is);
                return new StreamWorkbook(workbookReader);
            } else if (excelType == ExcelType.XLS) {
                return new HSSFWorkbook(is);
            } else {
                throw new NotSupportedException("Only support for .XLSX and .XLS!");
            }
        }

        /**
         * @return org.apache.poi.ss.usermodel.Workbook
         * read a giver {@link File} and return a new instance of {@link Workbook}.
         * recommended
         * @params file
         */
        public Workbook open(File file) throws IOException {
            if (file.getName().toUpperCase().endsWith(".XLSX")) {
                // using the user model of XSSF to compare with this project
                //return new XSSFWorkbook(file);

                WorkbookReader workbookReader = new WorkbookReader(this);
                workbookReader.init(file);
                return new StreamWorkbook(workbookReader);
            } else if (file.getName().toUpperCase().endsWith(".XLS")) {
                return new HSSFWorkbook(new FileInputStream(file));
            } else {
                throw new NotSupportedException("Only support for .XLSX and .XLS!");
            }
        }

        public int getRowCacheSize() {
            return rowCacheSize;
        }

        public Builder rowCacheSize(int rowCacheSize) {
            this.rowCacheSize = rowCacheSize;
            return this;
        }

        public int getBufferSize() {
            return bufferSize;
        }

        public Builder bufferSize(int bufferSize) {
            this.bufferSize = bufferSize;
            return this;
        }

        /**
         * If less than 0, no cache will be used and the entire table will be loaded into memory.
         *
         * @return the size of SharedString table cache
         */
        public int getSstCacheSize() {
            return sstCacheSize;
        }

        public Builder sstCacheSize(int sstCacheSize) {
            this.sstCacheSize = sstCacheSize;
            return this;
        }

        public int getSheetIndex() {
            return sheetIndex;
        }

        public Builder sheetIndex(int sheetIndex) {
            this.sheetIndex = sheetIndex;
            return this;
        }

        public String getPassword() {
            return password;
        }

        public Builder password(String password) {
            this.password = password;
            return this;
        }
    }
}
