package com.unaware.poi.excel.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.*;
import java.util.stream.Collectors;

/**
 * EasyRefiner-step1
 *
 * @author mno
 * @Modifier unaware
 * @date 2018/6/12 14:32
 */
public class AvailableInfoUtils {

    /**
     * 初始化可用信息
     *
     * @param style 单元格 风格
     * @return 信息
     */
    public static Map<OutputField, Object> initStyle(CellStyle style) {
        Map<OutputField, Object> map = new HashMap<>(OutputField.SORTED_FIELDS.size());
        map.put(OutputField.BorderTop, (int) style.getBorderTopEnum().getCode());
        map.put(OutputField.BorderBottom, (int) style.getBorderBottomEnum().getCode());
        map.put(OutputField.BorderLeft, (int) style.getBorderLeftEnum().getCode());
        map.put(OutputField.BorderRight, (int) style.getBorderRightEnum().getCode());
        map.put(OutputField.DataFormat, (int) style.getDataFormat());
        /*
          the values of FillBackgroundColor and FillForegroundColor are unsigned number
          But there is no unsigned type in Java, so Java treats it as a signed number
         */
        map.put(OutputField.FillBackgroundColor, ColorInfo.getColor(style.getFillBackgroundColorColor()).getRGB());
        map.put(OutputField.FillForegroundColor, ColorInfo.getColor(style.getFillForegroundColorColor()).getRGB());
        return map;
    }

    /**
     * 初始化可用信息
     *
     * @return 信息
     */
    public static Map<OutputField, Object> initDefault(int row, int column) {
        Map<OutputField, Object> map = OutputField.SORTED_FIELDS.stream().collect(Collectors.toMap(f -> f, f -> -1));
        fillingIndexAndMergeInfo(map, row, column, new ArrayList<>(0));
        return map;
    }

    /**
     * 初始化可用信息
     * Proposal for FontName: Collect all the fonts and put them into a Font Map.
     * For the Font of every cell, only save its index which is looked up in the Font Map by its name.
     *
     * @param map  已有的信息
     * @param font 字体信息
     */
    public static void initFont(Map<OutputField, Object> map, Font font) {
        map.put(OutputField.FontName, font.getFontName());
        map.put(OutputField.FontBold, font.getBold() ? 1 : 0);
    }

    /**
     * determine whether (row, column) is the beginning cell in the merged cell area
     *
     * @param row
     * @param column
     * @param MergedCells
     * @return
     */
    public static int isMergedBegin(int row, int column, List<MergedCell> MergedCells) {
        Optional<MergedCell> optional = MergedCells.stream().filter(v -> v.isMergedBegin(row, column)).findFirst();
        // if this cell is not in the merged cell area, this cell's value is -1
        return optional.map(MergedCell::getMergedId).orElse(-1);
    }

    /**
     * determine whether (row, column) is in the merged cell area
     *
     * @param row
     * @param column
     * @param MergedCells
     * @return
     */
    public static int getMergedIndex(int row, int column, List<MergedCell> MergedCells) {
        Optional<MergedCell> optional = MergedCells.stream().filter(v -> v.containsRow(row) && v.containsColumn(column)).findFirst();
        // if this cell is not in the merged cell area, this cell's value is -1
        return optional.map(MergedCell::getMergedId).orElse(-1);
    }

    /**
     * 装填行列标识以及合并信息
     *
     * @param map         信息
     * @param row         第几行 （0-based physical & logical）
     * @param column      第几列 （0-based physical & logical）
     * @param mergedCells 合并信息
     */
    public static boolean fillingIndexAndMergeInfo(Map<OutputField, Object> map, int row, int column, List<MergedCell> mergedCells) {
        map.put(OutputField.RowLabel, row);
        map.put(OutputField.ColumnLabel, column);
        int mergedId = getMergedIndex(row, column, mergedCells);
        map.put(OutputField.MergerCell, mergedId);
        return mergedId != -1;
    }

    /**
     * 按照一定顺序,写出来
     * 行标识、列标识 自动加一
     *
     * @param map 信息
     * @return list顺序的信息
     */
    public static List<String> toList(Map<OutputField, Object> map) {
        List<String> list = new ArrayList<>(map.size());
        OutputField.SORTED_FIELDS.forEach(field -> {
            if (map.containsKey(field)) {
                if (field == OutputField.RowLabel || field == OutputField.ColumnLabel) {
                    list.add(String.valueOf((int) map.get(field) + 1));
                } else {
                    list.add(String.valueOf(map.get(field)));
                }
            } else {
                list.add("");
            }
        });
        return list;
    }
}
