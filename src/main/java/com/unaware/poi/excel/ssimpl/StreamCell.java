package com.unaware.poi.excel.ssimpl;

import com.unaware.poi.excel.exception.NotSupportedException;
import com.unaware.poi.excel.util.DataUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Pattern;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * @author Unaware
 * @Title: StreamCell
 * @ProjectName excel
 * @Description: High level representation of a cell in a row of a spreadsheet.
 *                Cells can be numeric, formula-based or string-based (text). The cell type specifies this. String cells cannot conatin numbers and numeric cells cannot contain strings (at least according to our model). Client apps should do the conversions themselves. Formula cells have the formula string, as well as the formula result, which can be numeric or string.
 *                Cells should have their number (0 based) before being added to a row.
 *                It's based on org.apache.poi.ss.usermodel.Cell, override the methods of poi. Therefore you could use this like poi.
 *                There is only part of the reading method implemented
 * @date 2018/7/12 15:22
 */
public class StreamCell implements Cell {
    private static final String FALSE_AS_STRING = "0";
    private static final String TRUE_AS_STRING = "1";

    private final int colIndex;
    private final int rowIndex;
    private final boolean use1904Dates;

    private String formula;
    private String cachedFormulaResultType;
    private Row row;
    private CellStyle cellStyle;

    /**
     * The type of the cell data
     */
    private String type;
    private String numericFormat;
    private Short numericFormatIndex;
    private String rawContents;

    public StreamCell(int colIndex, int rowIndex, boolean use1904Dates) {
        this.colIndex = colIndex;
        this.rowIndex = rowIndex;
        this.use1904Dates = use1904Dates;
    }

    @Override
    public int getColumnIndex() {
        return colIndex;
    }

    @Override
    public int getRowIndex() {
        return rowIndex;
    }

    @Override
    public Sheet getSheet() {
        throw new NotSupportedException();
    }

    @Override
    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    /**
     * @param i
     * @deprecated
     */
    @Override
    public void setCellType(int i) {
        throw new NotSupportedException();
    }

    @Override
    public void setCellType(CellType cellType) {
        throw new NotSupportedException();
    }

    /**
     * @deprecated
     */
    @Override
    public int getCellType() {
        return getCellTypeEnum().getCode();
    }

    /**
     * determine the CellType based on the parameter "type"
     * @param type
     * @return
     */
    private  CellType judgeType(String type) {
        switch (type) {
            case "n":
                return CellType.NUMERIC;
            case "s":
            case "inlineStr":
                return STRING;
            case "str":
                return CellType.FORMULA;
            case "b":
                return CellType.BOOLEAN;
            case "e":
                return CellType.ERROR;
            default:
                throw new UnsupportedOperationException("Unsupported cell type '" + type + "'");
        }
    }

    @Override
    public CellType getCellTypeEnum() {
        if (rawContents == null || rawContents.length() == 0 || type == null) {
            return CellType.BLANK;
        } else {
            return judgeType(type);
        }
    }

    /**
     * @deprecated
     */
    @Override
    public int getCachedFormulaResultType() {
        return getCachedFormulaResultTypeEnum().getCode();
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum() {
        if (type != null && "str".equals(type)) {
            if (rawContents == null || cachedFormulaResultType == null) {
                return CellType.BLANK;
            } else {
                return judgeType(cachedFormulaResultType);
            }
        } else {
            throw new IllegalStateException("Only formula cells have cached results");
        }
    }

    @Override
    public void setCellValue(double v) {
        this.setRawContents(String.valueOf(v));
    }

    @Override
    public void setCellValue(Date date) {
        throw new NotSupportedException();
    }

    @Override
    public void setCellValue(Calendar calendar) {
        throw new NotSupportedException();
    }

    @Override
    public void setCellValue(RichTextString richTextString) {
        throw new NotSupportedException();
    }

    @Override
    public void setCellValue(String s) {
        this.setRawContents(s);
    }

    @Override
    public void setCellFormula(String s) throws FormulaParseException {
        throw new NotSupportedException();
    }

    @Override
    public String getCellFormula() {
        if (getCellTypeEnum() != CellType.FORMULA)
            throw typeMismatch("FORMULA", getCellTypeEnum().name(), false);
        return formula;
    }

    /**
     * Get the value of the cell as a number. For strings we throw an exception. For
     * blank cells we return a 0.
     *
     * @return the value of the cell as a number
     * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
     */
    @Override
    public double getNumericCellValue() {
        CellType cellType = getCellTypeEnum();
        switch(cellType) {
            case BLANK:
                return 0.0;
            case FORMULA:
                // fall-through
            case NUMERIC:
                if(rawContents == null || rawContents.isEmpty()) {
                    return 0.0;
                }
                try {
                    BigDecimal bd = new BigDecimal(rawContents);
                    return bd.doubleValue();
                } catch(NumberFormatException e) {
                        throw typeMismatch("NUMERIC", CellType.STRING.name(), false);
                }
            default:
                throw typeMismatch("NUMERIC", getCellTypeEnum().name(), false);
        }
    }

    /**
     * Get the value of the cell as a date.
     * For strings we throw an exception. For blank cells we return a null.
     *
     * @return the value of the cell as a date
     * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
     * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
     */
    @Override
    public Date getDateCellValue() {
        if (getCellTypeEnum() == STRING) {
            throw typeMismatch("DATE", "STRING", false);
        }
        if (getCellTypeEnum() == BLANK) {
            return null;
        }
        return rawContents == null ? null : HSSFDateUtil.getJavaDate(getNumericCellValue(), use1904Dates);
    }

    @Override
    public RichTextString getRichStringCellValue() {
        CellType cellType = getCellTypeEnum();
        XSSFRichTextString rt;
        switch (cellType) {
            case BLANK:
                rt = new XSSFRichTextString("");
                break;
            case STRING:
                rt = new XSSFRichTextString(getStringCellValue());
                break;
            default:
                throw new NotSupportedException();
        }
        return rt;
    }

    /**
     * Get the value of the cell as a string.
     * For numeric cells, we convert it to String and return the result.
     * For blank cells, we return an empty String.
     * For date cells, we return an formatted String
     * For formula cells, we return the calculation results
     * @return the value of the cell as a string
     */
    @Override
    public String getStringCellValue() {
        /*
          Determine whether the CellType is NUMERIC.
          Excel turns the date into NUMERIC to store
         */
        if (this.getCellTypeEnum() == NUMERIC) {
            /*
              If it is the date of the Chinese type, convert its format to "yyyy-MM-dd".
              yyyy-MM-dd: dataFormat=14,dataFormatString=m/d/yy
              yyyy"年"m"月"d"日": dataFormat=31/30,dataFormatString=reserved-0x1F
              上午/下午 hh"时"mm"分": dataFormat=55,dataFormatString=null
              HH:mm: dataFormat=20,dataFormatString=HH:mm
              hh"时"mm"分": dataFormat=32,dataFormatString=reserved-0x20
              hh"时"mm"分"ss"秒": dataFormat=33,dataFormatString=reserved-0x21
              hh"时"mm"分"ss"秒": dataFormat=56,dataFormatString=null
              yyyy"年"m"月": dataFormat=57,dataFormatString=null
              m"月"d"日": dataFormat=58,dataFormatString=reserved-0x1C
             */
            Date d;
            short dfIndex = this.getCellStyle().getDataFormat();
            if (dfIndex == 31 || dfIndex ==14 || dfIndex == 30) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue(), use1904Dates);
                return new SimpleDateFormat("yyyy-MM-dd").format(d);
            } else if (dfIndex == 55 || dfIndex == 32 || dfIndex == 20) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue(), use1904Dates);
                return new SimpleDateFormat("HH:mm").format(d);
            } else if (dfIndex == 56 || dfIndex ==33) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue(), use1904Dates);
                return new SimpleDateFormat("HH:mm:ss").format(d);
            } else if (dfIndex == 57) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue(), use1904Dates);
                return new SimpleDateFormat("yyyy-MM").format(d);
            } else if (dfIndex == 58) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue(), use1904Dates);
                return new SimpleDateFormat("MM-dd").format(d);
            }
            //Check if a cell contains a date
            if (DataUtil.isDateFormat(this)) {
                d = HSSFDateUtil.getJavaDate(this.getNumericCellValue());
                return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(d);
            }
            //return NUMERIC
            return String.valueOf(this.getNumericCellValue());
        }

        /*
          convert others' format to String
         */
        return rawContents == null ? "" : rawContents;
    }

    @Override
    public void setCellValue(boolean b) {
        throw new NotSupportedException();
    }

    @Override
    public void setCellErrorValue(byte b) {
        throw new NotSupportedException();
    }

    @Override
    public boolean getBooleanCellValue() {
        CellType cellType = this.getCellTypeEnum();
        switch(cellType) {
            case BLANK:
                return false;
            case BOOLEAN:
                return TRUE_AS_STRING.equals(getStringCellValue());
            case FORMULA:
                //YK: should throw an exception if requesting boolean value from a non-boolean formula
                return TRUE_AS_STRING.equals(getStringCellValue());
            default:
                throw typeMismatch("BOOLEAN", cellType.name(), false);
        }
    }

    @Override
    public byte getErrorCellValue() {
        CellType cellType = this.getCellTypeEnum();
        if(cellType == CellType.BLANK) {
            return 0;
        }else if(cellType == CellType.FORMULA){
                if(!isNumber(rawContents))
                    return FormulaError.forString(rawContents).getCode();
                else
                    return -1;
        } else if (cellType != CellType.ERROR) {
            throw typeMismatch("ERROR", cellType.name(), false);
        } else {
            return FormulaError.forString(rawContents).getCode();
        }
    }

    /**
     *  determine whether a string is numeric using a regular expression
     * @param str
     * @return
     */
    private static boolean isNumber(String str) {
        boolean isInt = Pattern.compile("^-?[1-9]\\d*$").matcher(str).find();
        boolean isDouble = Pattern.compile("^[-+]?([1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*|0?\\.0+|0)$").matcher(str).find();
        return isInt || isDouble;
    }

    /**
     * throw a RuntimeException to display the problem about format mismatched
     * @param expectedType
     * @param actualType
     * @param isFormulaCell
     * @return
     */
    private static RuntimeException typeMismatch(String expectedType, String actualType, boolean isFormulaCell) {
        String msg = "Cannot get a " + expectedType + " value from a " + actualType + " " + (isFormulaCell ? "formula " : "") + "cell";
        return new IllegalStateException(msg);
    }

    @Override
    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    @Override
    public CellStyle getCellStyle() {
        return this.cellStyle;
    }

    @Override
    public void setAsActiveCell() {
        throw new NotSupportedException();
    }

    @Override
    public CellAddress getAddress() {
        return new CellAddress(rowIndex, colIndex);
    }

    @Override
    public void setCellComment(Comment comment) {
        throw new NotSupportedException();
    }

    @Override
    public Comment getCellComment() {
        throw new NotSupportedException();
    }

    @Override
    public void removeCellComment() {
        throw new NotSupportedException();
    }

    @Override
    public Hyperlink getHyperlink() {
        throw new NotSupportedException();
    }

    @Override
    public void setHyperlink(Hyperlink hyperlink) {
        throw new NotSupportedException();
    }

    @Override
    public void removeHyperlink() {
        throw new NotSupportedException();
    }

    @Override
    public CellRangeAddress getArrayFormulaRange() {
        throw new NotSupportedException();
    }

    @Override
    public boolean isPartOfArrayFormulaGroup() {
        throw new NotSupportedException();
    }

    public String getType() {
        return this.type;
    }

    public String getNumericFormat() {
        return this.numericFormat;
    }

    public Short getNumericFormatIndex() {
        return this.numericFormatIndex;
    }

    public void setNumericFormatIndex(Short dataFormat) {
        this.numericFormatIndex = dataFormat;
    }

    public void setNumericFormat(String formatString) {
        this.numericFormat = formatString;
    }

    public void setRawContents(String contents) {
        this.rawContents = contents;
    }

    public void setType(String value) {
        if ("str".equals(value)) {
            // this is a formula cell, cache the value's type
            cachedFormulaResultType = this.type;
        }
        this.type = value;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }
}
