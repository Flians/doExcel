package com.unaware.poi.excel.util;

import com.google.common.collect.Range;

import java.io.File;
import java.io.Serializable;

public class MixedFile implements Serializable {
    private static final long serialVersionUID = 1705493650770372671L;

    /**
     * 原始文件
     */
    private File original;

    /**
     * 合并信息文件（csv文件没有,Excel 才有）
     */
    private File merge;

    /**
     * sheet 名字（csv 没有）
     */
    private String sheetName = "";

    /**
     * 有效行, (1-base)
     */
    private Range<Integer> validRows;

    /**
     * 有效列, (1-base)
     */
    private Range<Integer> validCols;


    public MixedFile() {
    }

    public MixedFile(File original) {
        this(original, null);
    }

    public MixedFile(File original, File merge) {
        this(original, merge, "");
    }

    public MixedFile(File original, File merge, String sheetName) {
        this.original = original;
        this.merge = merge;
        this.sheetName = sheetName;
    }

    public File getOriginal() {
        return original;
    }

    public void setOriginal(File original) {
        this.original = original;
    }

    public File getMerge() {
        return merge;
    }

    public void setMerge(File merge) {
        this.merge = merge;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Range<Integer> getValidRows() {
        return validRows;
    }

    public void setValidRows(Range<Integer> validRows) {
        this.validRows = validRows;
    }

    public Range<Integer> getValidCols() {
        return validCols;
    }

    public void setValidCols(Range<Integer> validCols) {
        this.validCols = validCols;
    }

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder("MixedFile{");
        sb.append("original=").append(original);
        sb.append(", merge=").append(merge);
        sb.append(", sheetName='").append(sheetName).append('\'');
        sb.append(", validRows=").append(validRows);
        sb.append(",  validCols=").append(validCols);
        sb.append('}');
        return sb.toString();
    }
}
