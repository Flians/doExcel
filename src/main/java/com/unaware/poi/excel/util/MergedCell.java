package com.unaware.poi.excel.util;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author Unaware
 * @Description: redefine the CellRangeAddress.
 *                 add attributes: mergedId and value
 * @Title: MergedCell
 * @ProjectName step1
 * @date 2018/8/3 10:40
 */
public class MergedCell extends CellRangeAddress {
    /**
     * the index of this MergedCell in  List<CellRangeAddress>
     */
    private int mergedId;
    private String value;

    public MergedCell(int firstRow, int lastRow, int firstCol, int lastCol) {
        super(firstRow, lastRow, firstCol, lastCol);
        if (lastRow < firstRow || lastCol < firstCol) {
            throw new IllegalArgumentException("Invalid cell range, having lastRow < firstRow || lastCol < firstCol, had rows " + lastRow + " >= " + firstRow + " or cells " + lastCol + " >= " + firstCol);
        }
    }

    public MergedCell(CellRangeAddress rangeAddress) {
        this(rangeAddress.getFirstRow(), rangeAddress.getLastRow(), rangeAddress.getFirstColumn(), rangeAddress.getLastColumn());
    }

    public MergedCell(int mergedId, String value, int firstRow, int lastRow, int firstCol, int lastCol) {
        super(firstRow, lastRow, firstCol, lastCol);
        this.mergedId = mergedId;
        this.value = value;
        if (lastRow < firstRow || lastCol < firstCol) {
            throw new IllegalArgumentException("Invalid cell range, having lastRow < firstRow || lastCol < firstCol, had rows " + lastRow + " >= " + firstRow + " or cells " + lastCol + " >= " + firstCol);
        }
    }

    public MergedCell(int mergedId, String value, CellRangeAddress rangeAddress){
        super(rangeAddress.getFirstRow(), rangeAddress.getLastRow(), rangeAddress.getFirstColumn(), rangeAddress.getLastColumn());
        this.mergedId = mergedId;
        this.value = value;
    }

    public MergedCell(int mergedId, CellRangeAddress rangeAddress) {
        super(rangeAddress.getFirstRow(), rangeAddress.getLastRow(), rangeAddress.getFirstColumn(), rangeAddress.getLastColumn());
        this.mergedId = mergedId;
        this.value = "";
    }

    public boolean isMergedBegin(int row, int column) {
        return getFirstRow() == row && getFirstColumn() == column;
    }

    public int getMergedId() {
        return mergedId;
    }

    public void setMergedId(int mergedId) {
        this.mergedId = mergedId;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
