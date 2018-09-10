package com.unaware.poi.excel;

import com.unaware.poi.excel.streamreader.StreamReader;
import com.unaware.poi.excel.util.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Pattern;

/**
 * @author Unaware
 * @Description: write the data and style of .xls/.xlsx File into the .csv file
 * @Title: SSConverter
 * @ProjectName step1
 * @date 2018/7/31 11:23
 */
public class SSConverter implements AutoCloseable {

    /**
     * Filter line symbol "\r\n"
     */
    private static final Pattern CELL_FILTER = Pattern.compile("[\r\n\"]");

    /**
     * the Excel file
     */
    private File file;

    /**
     * parse xls和xlsx
     */
    private Workbook workbook;

    /**
     * determine whether obtaining the style information
     */
    private boolean enableAvailableInfo = false;

    /**
     * the information of .csv file
     */
    private List<MixedFile> mixedFiles;

    public void path(File file, int sstCacheSize, int rowCacheSize, int sheetIndex) throws Exception {
        this.file = file;
        this.workbook = StreamReader.builder().sstCacheSize(sstCacheSize).rowCacheSize(rowCacheSize).sheetIndex(sheetIndex).open(file);
        this.mixedFiles = writeIntoCSV();
        this.close();
    }

    /**
     * use default parameters
     *
     * @param file
     * @throws IOException
     */
    public void path(File file) throws Exception {
        this.file = file;
        this.workbook = StreamReader.builder().open(file);
        this.mixedFiles = writeIntoCSV();
        this.close();
    }

    /**
     * Filter line symbol "\r\n" in cell data
     *
     * @param cellValue
     * @return
     */
    private String filterCell(String cellValue) {
        if (cellValue == null) {
            return "";
        }
        return CELL_FILTER.matcher(cellValue).replaceAll("");
    }


    public File getFile() {
        return file;
    }


    /**
     * determine whether obtaining the style information
     *
     * @param enable true: obtain
     */
    public void enableAvailableInfo(boolean enable) {
        enableAvailableInfo = enable;
    }

    /**
     * Get the number of the sheet the beginning and the end
     *
     * @return
     */
    private int availableInfoRowNum() {
        return 200;
    }

    /**
     * get the value in top left corner and the merge area, and assign this value to every cell in the merged cell area.
     *
     * @param sheet
     * @return
     */
    private List<MergedCell> handleMergedRegion(Sheet sheet) {
        int numMergedRegions = sheet.getNumMergedRegions();
        List<MergedCell> MergedCells = new ArrayList<>(numMergedRegions);
        for (int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
            MergedCells.add(new MergedCell(i, cellRangeAddress));
        }
        return MergedCells;
    }

    /**
     * @return original file and available information of file
     */
    public List<MixedFile> getMixedFiles() {
        return this.mixedFiles;
    }

    /**
     * respectively write the contents of cells and style of cells into the .csv file
     *
     * @return original file and available information of file
     */
    private List<MixedFile> writeIntoCSV() {
        int sheetNum = workbook.getNumberOfSheets();
        if (sheetNum == 0) {
            return new ArrayList<>(0);
        }
        List<MixedFile> mixedFiles = new ArrayList<>(sheetNum);
        workbook.forEach(sheet -> {
            // handle mergedCell
            List<MergedCell> mergedCells = handleMergedRegion(sheet);
            // the out file generated in random
            File originalFile = new File("src\\test\\resources\\output\\" + DataUtil.getUUID() + ".csv");
            File availableInfoFile = new File("src\\test\\resources\\output\\" + DataUtil.getUUID() + ".csv");
            try (
                    CsvWriter writerOriginal = CsvWriter.utf8(originalFile);
                    CsvWriter writerAvailableInfo = CsvWriter.utf8(availableInfoFile)) {
                int ltNum = availableInfoRowNum(), rtNum = sheet.getLastRowNum() - availableInfoRowNum();
                rtNum = rtNum < ltNum ? Integer.MAX_VALUE : rtNum;
                int finalRtNum = rtNum;
                sheet.forEach(row -> {
                    List<String> originalData = new ArrayList<>(32);
                    AtomicInteger col = new AtomicInteger();
                    row.forEach(c -> {
                        int mergedIndex = AvailableInfoUtils.isMergedBegin(row.getRowNum(), col.get(), mergedCells);
                        if(mergedIndex != -1){
                            mergedCells.get(mergedIndex).setValue(filterCell(DataUtil.getCellValue(c)));
                        }
                        CellStyle style = null;
                        if (c == null) {
                            originalData.add("");
                        } else {
                            // blank cell
                            int ss = c.getColumnIndex();
                            while (originalData.size() < ss) {
                                originalData.add("");
                                // As long as there is a combination of information, write its style information
                                if (enableAvailableInfo && (row.getRowNum() < ltNum || row.getRowNum() > finalRtNum)) {
                                    Map<OutputField, Object> map = AvailableInfoUtils.initDefault(row.getRowNum(), col.get());
                                    if (AvailableInfoUtils.fillingIndexAndMergeInfo(map, row.getRowNum(), col.get(), mergedCells)) {
                                        writerAvailableInfo.write(AvailableInfoUtils.toList(map));
                                    }
                                }
                                col.getAndIncrement();
                            }
                            mergedIndex = AvailableInfoUtils.getMergedIndex(row.getRowNum(), col.get(), mergedCells);
                            originalData.add(mergedIndex == -1 ? filterCell(DataUtil.getCellValue(c)) : mergedCells.get(mergedIndex).getValue());
                            col.set(c.getColumnIndex());
                            style = c.getCellStyle();
                        }
                        // write the style information into .csv file
                        if (enableAvailableInfo && (row.getRowNum() < ltNum || row.getRowNum() > finalRtNum)) {
                            Map<OutputField, Object> map;
                            if (style != null) {
                                map = AvailableInfoUtils.initStyle(style);
                                Font font = workbook.getFontAt(style.getFontIndex());
                                AvailableInfoUtils.initFont(map, font);
                            } else {
                                map = AvailableInfoUtils.initDefault(row.getRowNum(), col.get());
                            }
                            // As long as it's a merged cell or it has the style , write its style information
                            if (AvailableInfoUtils.fillingIndexAndMergeInfo(map, row.getRowNum(), col.get(), mergedCells) || style != null) {
                                writerAvailableInfo.write(AvailableInfoUtils.toList(map));
                            }
                        }
                        col.getAndIncrement();
                    });
                    writerOriginal.write(originalData);
                });
                mixedFiles.add(new MixedFile(originalFile, availableInfoFile, sheet.getSheetName()));
            }
            mergedCells.clear();
        });
        return mixedFiles;
    }

    /**
     * Closes this resource, relinquishing any underlying resources.
     * This method is invoked automatically on objects managed by the
     * {@code try}-with-resources statement.
     * <p>
     * <p>While this interface method is declared to throw {@code
     * exception}, implementers are <em>strongly</em> encouraged to
     * declare concrete implementations of the {@code close} method to
     * throw more specific exceptions, or to throw no exception at all
     * if the close operation cannot fail.
     * <p>
     * <p> Cases where the close operation may fail require careful
     * attention by implementers. It is strongly advised to relinquish
     * the underlying resources and to internally <em>mark</em> the
     * resource as closed, prior to throwing the exception. The {@code
     * close} method is unlikely to be invoked more than once and so
     * this ensures that the resources are released in a timely manner.
     * Furthermore it reduces problems that could arise when the resource
     * wraps, or is wrapped, by another resource.
     * <p>
     * <p><em>Implementers of this interface are also strongly advised
     * to not have the {@code close} method throw {@link
     * InterruptedException}.</em>
     * <p>
     * This exception interacts with a thread's interrupted status,
     * and runtime misbehavior is likely to occur if an {@code
     * InterruptedException} is {@linkplain Throwable#addSuppressed
     * suppressed}.
     * <p>
     * More generally, if it would cause problems for an
     * exception to be suppressed, the {@code AutoCloseable.close}
     * method should not throw it.
     * <p>
     * <p>Note that unlike the {@link Closeable#close close}
     * method of {@link Closeable}, this {@code close} method
     * is <em>not</em> required to be idempotent.  In other words,
     * calling this {@code close} method more than once may have some
     * visible side effect, unlike {@code Closeable.close} which is
     * required to have no effect if called more than once.
     * <p>
     * However, implementers of this interface are strongly encouraged
     * to make their {@code close} methods idempotent.
     *
     * @throws Exception if this resource cannot be closed
     */
    @Override
    public void close() throws Exception {
        if (workbook != null) {
            workbook.close();
            workbook = null;
        }
    }
}
