package com.unaware.poi.excel.ssimpl;

import com.unaware.poi.excel.streamreader.SheetReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author Unaware
 * @Title: StreamSheet
 * @ProjectName excel
 * @Description: High level representation of a Excel worksheet.
 *                Sheets are the central structures within a workbook, and are where a user does most of his spreadsheet work.
 *                The most common type of sheet is the worksheet, which is represented as a grid of cells.
 *                Worksheet cells can contain text, numbers, dates, and formulas. Cells can also be formatted.
 *                It's based on org.apache.poi.ss.usermodel.Sheet, override the methods of poi. Therefore you could use this like poi.
 *                There is only part of the reading method implemented
 * @date 2018/7/12 15:21
 */
public class StreamSheet implements Sheet {
    private final String name;
    private final SheetReader sheetReader;

    /**
     * constructor
     * @param name
     * @param sheetReader
     */
    public StreamSheet(String name, SheetReader sheetReader) {
        this.name = name;
        this.sheetReader = sheetReader;
    }

    @Override
    public Row createRow(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeRow(Row row) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Row getRow(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getPhysicalNumberOfRows() {
        return this.getLastRowNum() - this.getFirstRowNum() + 1;
    }

    @Override
    public int getFirstRowNum() {
        return sheetReader.getFirstRowNum();
    }

    @Override
    public int getLastRowNum() {
        return sheetReader.getLastRowNum();
    }

    @Override
    public void setColumnHidden(int i, boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isColumnHidden(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRightToLeft(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isRightToLeft() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setColumnWidth(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getColumnWidth(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public float getColumnWidthInPixels(int i) {
        return sheetReader.getColWidth().get(i).floatValue();
    }

    @Override
    public void setDefaultColumnWidth(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getDefaultColumnWidth() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getDefaultRowHeight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public float getDefaultRowHeightInPoints() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDefaultRowHeight(short i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDefaultRowHeightInPoints(float v) {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle getColumnStyle(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addMergedRegion(CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int addMergedRegionUnsafe(CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void validateMergedRegions() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setVerticallyCenter(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHorizontallyCenter(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getHorizontallyCenter() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getVerticallyCenter() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeMergedRegion(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeMergedRegions(Collection<Integer> collection) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getNumMergedRegions() {
        return sheetReader.getNumMergedRegions();
    }

    @Override
    public CellRangeAddress getMergedRegion(int i) {
        return sheetReader.getMergedRegions().get(i);
    }

    @Override
    public List<CellRangeAddress> getMergedRegions() {
        return sheetReader.getMergedRegions();
    }

    @Override
    public Iterator<Row> rowIterator() {
        return sheetReader.iterator();
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
    public void setAutobreaks(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDisplayGuts(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDisplayZeros(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isDisplayZeros() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFitToPage(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowSumsBelow(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowSumsRight(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getAutobreaks() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getDisplayGuts() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getFitToPage() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getRowSumsBelow() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getRowSumsRight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isPrintGridlines() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintGridlines(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isPrintRowAndColumnHeadings() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setPrintRowAndColumnHeadings(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public PrintSetup getPrintSetup() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Header getHeader() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Footer getFooter() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setSelected(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public double getMargin(short i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setMargin(short i, double v) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getProtect() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void protectSheet(String s) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getScenarioProtect() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setZoom(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getTopRow() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getLeftCol() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void showInPane(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void shiftRows(int i, int i1, int i2) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void shiftRows(int i, int i1, int i2, boolean b, boolean b1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void createFreezePane(int i, int i1, int i2, int i3) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void createFreezePane(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void createSplitPane(int i, int i1, int i2, int i3, int i4) {
        throw new UnsupportedOperationException();
    }

    @Override
    public PaneInformation getPaneInformation() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDisplayGridlines(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isDisplayGridlines() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDisplayFormulas(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isDisplayFormulas() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDisplayRowColHeadings(boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isDisplayRowColHeadings() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowBreak(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isRowBroken(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeRowBreak(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int[] getRowBreaks() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int[] getColumnBreaks() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setColumnBreak(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isColumnBroken(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeColumnBreak(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setColumnGroupCollapsed(int i, boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void groupColumn(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void ungroupColumn(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void groupRow(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void ungroupRow(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowGroupCollapsed(int i, boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setDefaultColumnStyle(int i, CellStyle cellStyle) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void autoSizeColumn(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void autoSizeColumn(int i, boolean b) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Comment getCellComment(CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Map<CellAddress, ? extends Comment> getCellComments() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Drawing<?> getDrawingPatriarch() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Drawing<?> createDrawingPatriarch() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Workbook getWorkbook() {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getSheetName() {
        return name;
    }

    @Override
    public boolean isSelected() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellRange<? extends Cell> setArrayFormula(String s, CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellRange<? extends Cell> removeArrayFormula(Cell cell) {
        throw new UnsupportedOperationException();
    }

    @Override
    public DataValidationHelper getDataValidationHelper() {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends DataValidation> getDataValidations() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void addValidationData(DataValidation dataValidation) {
        throw new UnsupportedOperationException();
    }

    @Override
    public AutoFilter setAutoFilter(CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public SheetConditionalFormatting getSheetConditionalFormatting() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellRangeAddress getRepeatingRows() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellRangeAddress getRepeatingColumns() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRepeatingRows(CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRepeatingColumns(CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getColumnOutlineLevel(int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Hyperlink getHyperlink(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Hyperlink getHyperlink(CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    @Override
    public List<? extends Hyperlink> getHyperlinkList() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellAddress getActiveCell() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setActiveCell(CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Row> iterator() {
        return sheetReader.iterator();
    }

    public SheetReader getReader() {
        return sheetReader;
    }
}
