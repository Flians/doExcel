package com.unaware.poi.excel.util;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * @author Unaware
 * @Description: ${description}
 * @Title: CsvWriter
 * @ProjectName doExcel
 * @date 2018/9/11 1:30
 */
public final class CsvWriter implements AutoCloseable {
    private static final Logger LOGGER = LoggerFactory.getLogger(CsvWriter.class);
    private CSVPrinter printer;

    private CsvWriter(Writer writer) {
        try {
            this.printer = CSVFormat.INFORMIX_UNLOAD_CSV.withIgnoreEmptyLines(false).withAllowMissingColumnNames().withRecordSeparator("\r\n").print(writer);
        } catch (IOException var3) {
            LOGGER.error("发生了错误:{}", var3.getMessage());
            throw new IllegalArgumentException(var3.getMessage());
        }
    }

    public static CsvWriter utf8(File file) {
        return build(file, "UTF-8");
    }

    public static CsvWriter gbk(File file) {
        return build(file, "gbk");
    }

    public static CsvWriter build(File file, String charsetName) {
        OutputStreamWriter writer;
        try {
            writer = new OutputStreamWriter(new FileOutputStream(file), charsetName);
        } catch (IOException var4) {
            throw new IllegalArgumentException("文件操作失败");
        }

        return new CsvWriter(writer);
    }

    public boolean write(Iterable<String> values) {
        boolean flag = true;

        try {
            this.printer.printRecord(values);
        } catch (IOException var4) {
            flag = false;
            LOGGER.error("写一行数据,发生了错误:{}", var4.getMessage());
        }

        return flag;
    }

    public void close() {
        if (this.printer != null) {
            try {
                this.printer.close();
            } catch (IOException var2) {
                LOGGER.error("DataWriter 关流发生了错误:{}", var2.getMessage());
            }
        }

    }
}
