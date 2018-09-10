package com.unaware.poi.excel.ssimpl;

import com.unaware.poi.excel.exception.NotSupportedException;
import org.apache.poi.ss.usermodel.*;

import java.util.Iterator;
import java.util.TreeMap;

/**
 * @author Unaware
 * @Title: StreamRow
 * @ProjectName excel
 * @Description: High level representation of a row of a spreadsheet.
 *                It's based on org.apache.poi.ss.usermodel.Row, override the methods of poi. Therefore you could use this like poi.
 *                There is only part of the reading method implemented
 * @date 2018/7/12 15:22
 */
public class StreamRow implements Row {
    private int rowIndex;
    private TreeMap<Integer, Cell> cellMap;

    public StreamRow(int rowIndex) {
        this.rowIndex = rowIndex;
        this.cellMap = new TreeMap<>();
    }

    public TreeMap<Integer, Cell> getCellMap() {
        return this.cellMap;
    }

    @Override
    public Cell createCell(int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * @param i
     * @param i1
     * @deprecated
     */
    @Override
    public Cell createCell(int i, int i1) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Cell createCell(int i, CellType cellType) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeCell(Cell cell) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowNum(int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * Get row number this row represents
     * @return the row number (0 based)
     */
    @Override
    public int getRowNum() {
        return this.rowIndex;
    }

    @Override
    public Cell getCell(int i) {
        return this.cellMap.get(i);
    }

    @Override
    public Cell getCell(int i, MissingCellPolicy missingCellPolicy) {
        StreamCell cell = (StreamCell)this.cellMap.get(i);
        if (missingCellPolicy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
            if (cell == null) {
                return new StreamCell(i, this.rowIndex, false);
            }
        }else if (missingCellPolicy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
                if(cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                    return null;
                }
        }
        return cell;
    }

    @Override
    public short getFirstCellNum() {
        if(this.cellMap.size() == 0){
            return  -1;
        }
        return this.cellMap.firstKey().shortValue();
    }

    @Override
    public short getLastCellNum() {
        return (short) (this.cellMap.size() == 0?-1:this.cellMap.lastEntry().getValue().getColumnIndex() + 1);
    }

    @Override
    public int getPhysicalNumberOfCells() {
        return cellMap.size();
    }

    @Override
    public void setHeight(short i) {
        throw new NotSupportedException();
    }

    @Override
    public void setZeroHeight(boolean b) {
        throw new NotSupportedException();
    }

    @Override
    public boolean getZeroHeight() {
        throw new NotSupportedException();
    }

    @Override
    public void setHeightInPoints(float v) {
        throw new NotSupportedException();
    }

    @Override
    public short getHeight() {
        throw new NotSupportedException();
    }

    @Override
    public float getHeightInPoints() {
        throw new NotSupportedException();
    }

    @Override
    public boolean isFormatted() {
        throw new NotSupportedException();
    }

    @Override
    public CellStyle getRowStyle() {
        throw new NotSupportedException();
    }

    @Override
    public void setRowStyle(CellStyle cellStyle) {
        throw new NotSupportedException();
    }

    @Override
    public Iterator<Cell> cellIterator() {
        return this.cellMap.values().iterator();
    }

    @Override
    public Sheet getSheet() {
        throw new NotSupportedException();
    }

    @Override
    public int getOutlineLevel() {
        throw new NotSupportedException();
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Cell> iterator() {
        return this.cellMap.values().iterator();
    }
}
