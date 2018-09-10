package com.unaware.poi.excel.ssimpl;

import com.unaware.poi.excel.exception.MissingSheetException;
import com.unaware.poi.excel.streamreader.WorkbookReader;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;

import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

/**
 * @author Unaware
 * @Title: StreamWorkbook
 * @ProjectName excel
 * @Description: High level representation of a Excel workbook. This is the first object most users will construct whether they are reading or writing a workbook..
 *  *             It's based on org.apache.poi.ss.usermodel.Workbook, override the methods of poi. Therefore you could use this like poi.
 *  *             There is only part of the reading method implemented
 * @date 2018/7/12 15:18
 */
public class StreamWorkbook implements Workbook, AutoCloseable {
    private  final WorkbookReader workbookReader;

    /**
     * constructor
     * @param workbookReader
     */
    public StreamWorkbook(WorkbookReader workbookReader) {
        this.workbookReader = workbookReader;
    }

    @Override
    public int getActiveSheetIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setActiveSheet(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getFirstVisibleTab() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFirstVisibleTab(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetOrder(String s, int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSelectedTab(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetName(int i, String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getSheetName(int i) {
        return workbookReader.getSheetProperties().get(i).get("name");
    }

    @Override
    public int getSheetIndex(String s) {
        return findSheetByName(s);
    }

    @Override
    public int getSheetIndex(Sheet sheet) {
        if(sheet instanceof StreamSheet) {
            return findSheetByName(sheet.getSheetName());
        } else {
            throw new UnsupportedOperationException("Cannot use non-StreamingSheet sheets");
        }
    }

    @Override
    public Sheet createSheet() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet createSheet(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet cloneSheet(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Sheet> sheetIterator() {
        return iterator();
    }

    @Override
    public int getNumberOfSheets() {
        return workbookReader.getSheetProperties().size();
    }

    @Override
    public Sheet getSheetAt(int i) {
        return workbookReader.getSheets().get(i);
    }

    @Override
    public Sheet getSheet(String s) {
        int index = findSheetByName(s);
        if (index == -1) {
            throw new MissingSheetException("Sheet '" + s + "' does not exist");
        }
        return workbookReader.getSheets().get(index);
    }

    @Override
    public void removeSheetAt(int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * Create a new Font and add it to the workbook's font table
     *
     * @return new font object
     */
    @Override
    public Font createFont() {
        XSSFFont font = new XSSFFont(CTFont.Factory.newInstance());
        font.registerTo(workbookReader.getStylesSource());
        return font;
    }

    @Override
    public Font findFont(boolean b, short i, short i1, String s, boolean b1, boolean b2, short i2, byte b3) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getNumberOfFonts() {
        return (short) workbookReader.getStylesSource().getFonts().size();
    }

    @Override
    public Font getFontAt(short i) {
        if(i < 0 || i >= workbookReader.getStylesSource().getFonts().size()){
            return null;
        } else {
            return workbookReader.getStylesSource().getFonts().get(i);
        }
    }

    @Override
    public CellStyle createCellStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNumCellStyles() {
        return workbookReader.getStylesSource().getNumCellStyles();
    }

    @Override
    public CellStyle getCellStyleAt(int i) {
        return workbookReader.getStylesSource().getStyleAt(i);
    }

    @Override
    public void write(OutputStream outputStream) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void close() {
        try {
            workbookReader.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public int getNumberOfNames() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name getName(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends Name> getNames(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends Name> getAllNames() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name getNameAt(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name createName() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNameIndex(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(Name name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int linkExternalWorkbook(String s, Workbook workbook) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintArea(int i, String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintArea(int i, int i1, int i2, int i3, int i4) {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getPrintArea(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removePrintArea(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Row.MissingCellPolicy getMissingCellPolicy() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {
        throw new UnsupportedOperationException();
    }

    @Override
    public DataFormat createDataFormat() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addPicture(byte[] bytes, int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends PictureData> getAllPictures() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CreationHelper getCreationHelper() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isHidden() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHidden(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isSheetHidden(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isSheetVeryHidden(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetHidden(int i, boolean b) {
        throw new UnsupportedOperationException();
    }

    /**
     * @param i
     * @param i1
     * @deprecated
     */
    @Override
    public void setSheetHidden(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public SheetVisibility getSheetVisibility(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSheetVisibility(int i, SheetVisibility sheetVisibility) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void addToolPack(UDFFinder udfFinder) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setForceFormulaRecalculation(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getForceFormulaRecalculation() {
        throw new UnsupportedOperationException();
    }

    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addOlePackage(byte[] bytes, String s, String s1, String s2) {
        throw new UnsupportedOperationException();
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Sheet> iterator() {
        return workbookReader.iterator();
    }

    private int findSheetByName(String name) {
        for(int i = 0; i < workbookReader.getSheetProperties().size(); i++) {
            if(workbookReader.getSheetProperties().get(i).get("name").equals(name)) {
                return i;
            }
        }
        return -1;
    }
}
