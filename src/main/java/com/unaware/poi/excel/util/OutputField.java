package com.unaware.poi.excel.util;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author Unaware
 * @Description: ${description}
 * @Title: OutputField
 * @ProjectName doExcel
 * @date 2018/9/11 1:18
 */
public enum OutputField {
        ColumnLabel("cell_x"),
        RowLabel("cell_y"),
        BorderTop("top_border"),
        BorderBottom("bottom_border"),
        BorderLeft("left_border"),
        BorderRight("right_border"),
        DataFormat("data_format"),
        FillBackgroundColor("back_color"),
        FillForegroundColor("fore_color"),
        FontName("font_name"),
        FontBold("font_bold"),
        MergerCell("merge_idx");

        public static final List<OutputField> SORTED_FIELDS = Collections.unmodifiableList((List)Arrays.stream(values()).sorted().collect(Collectors.toList()));
        private String field;

        private OutputField(String field) {
            this.field = field;
        }

        public String getField() {
            return this.field;
        }
    }
