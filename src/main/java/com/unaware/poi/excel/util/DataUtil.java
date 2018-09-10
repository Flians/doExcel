package com.unaware.poi.excel.util;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.FORMULA;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * @author Unaware
 * @Description: Provide some methods of data processing
 * @Title: DataUtil
 * @ProjectName step1
 * @date 2018/7/27 13:46
 */
public class DataUtil {
    private static final NumberFormat percent = NumberFormat.getPercentInstance();
    /**
     * yyyy-MM-dd: dataFormat=14,dataFormatString=m/d/yy
     * yyyy"年"m"月"d"日": dataFormat=31/30,dataFormatString=reserved-0x1F
     * 上午/下午 hh"时"mm"分": dataFormat=55,dataFormatString=null
     * HH:mm: dataFormat=20,dataFormatString=HH:mm
     * hh"时"mm"分": dataFormat=32,dataFormatString=reserved-0x20
     * hh"时"mm"分"ss"秒": dataFormat=33,dataFormatString=reserved-0x21
     * hh"时"mm"分"ss"秒": dataFormat=56,dataFormatString=null
     * yyyy"年"m"月": dataFormat=57,dataFormatString=null
     * m"月"d"日": dataFormat=58,dataFormatString=reserved-0x1C
     * <p>
     * 2  represents date + time
     * 1  represents date
     * 0  represents time
     * -1 unknown
     */
    private static final Map<Short, Integer> reservedMap;
    private static final Set<String> znMap;

    static {
        reservedMap = new HashMap<>();
        reservedMap.put((short) 20, 0);
        reservedMap.put((short) 32, 0);
        reservedMap.put((short) 33, 0);
        reservedMap.put((short) 55, 0);
        reservedMap.put((short) 56, 0);

        reservedMap.put((short) 14, 1);
        reservedMap.put((short) 30, 1);
        reservedMap.put((short) 31, 1);
        reservedMap.put((short) 57, 1);
        reservedMap.put((short) 58, 1);

        reservedMap.put((short) 22, 2);

        znMap = new HashSet<>();
        znMap.add("年");
        znMap.add("月");
        znMap.add("日");
        znMap.add("时");
        znMap.add("分");
        znMap.add("秒");
        znMap.add("上午");
        znMap.add("下午");
        znMap.add("aaa;");
        znMap.add("aaaa;");
        znMap.add("AM");
        znMap.add("PM");
    }

    /**
     * Determine the numeric type, automatically resolve date formats, and other special types.
     * The date judgement of POI is only applicable to the date format in Europe and America.
     * It does not support the Chinese date, and uses two ways(including isReserved and isDateFormat) to determine the date of the Chinese format.
     * For Date, return "yyyy-MM-dd", "HH:mm:ss" or "yyyy-MM-dd HH:mm:ss"
     * For Formula, return the calculation results
     * For blank, we return an empty String.
     * For numeric, we convert it to String and return the result.
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        if (cell == null)
            return "";
        try {
            /*
              Determine whether the CellType is NUMERIC.
              Excel will turn the date into NUMERIC to store
             */
            if (cell.getCellTypeEnum() == NUMERIC) {
                Date d = cell.getDateCellValue();
                String format = cell.getCellStyle().getDataFormatString();
                int tt;
                if ((tt = isReserved(cell.getCellStyle().getDataFormat())) != -1) {
                    switch (tt) {
                        case 0:
                            return new SimpleDateFormat("HH:mm:ss").format(d);
                        case 1:
                            return new SimpleDateFormat("yyyy-MM-dd").format(d);
                        case 2:
                            return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(d);
                    }
                } else if (format != null) {
                    /*
                      Remove the "[...]" in front of the string, because sometimes it contains the letter D.
                     */
                    format = format.replaceAll("^\\[.*]", "").toUpperCase();

                    if (format.matches("^((?![YD]).)*((H.*M)|(M.*S))((?![YD]).)*$")) {
                        return new SimpleDateFormat("HH:mm:ss").format(d);
                    } else if (format.matches("^((?![SH]).)*(M|AAAA;|AAA;)((?![SH]).)*$")) {
                        return new SimpleDateFormat("yyyy-MM-dd").format(d);
                    } else if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(d);
                    }
                }
                // handle some numeric cells.
                String temp = handleNumeric(cell);
                if(temp != null) {
                    return temp;
                }
            } else if (cell.getCellTypeEnum() == FORMULA) {
                try {
                    // handle some numeric cells.
                    String temp = handleNumeric(cell);
                    if(temp != null) {
                        return temp;
                    } else {
                        BigDecimal bigDecimal = new BigDecimal(String.valueOf(cell.getNumericCellValue())).setScale(14, BigDecimal.ROUND_HALF_UP);
                        return bigDecimal.stripTrailingZeros().toPlainString();
                    }
                } catch (IllegalStateException | NumberFormatException e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (IllegalStateException e1) {
                        return FormulaError.forInt(cell.getErrorCellValue()).getString();
                    }
                }
            }
            /*
              convert others' format to String, the value is determined by the actual data type
             */
            HSSFDataFormatter df = new HSSFDataFormatter();
            return df.formatCellValue(cell);
        } catch (Exception e) {
            System.out.println("ERROR: " + cell.getAddress().toString() + " -> " + e.getMessage());
            return "#ERROR";
        }
    }

    /**
     * Determine whether it is a date format reserved field
     *
     * @param reserved
     * @return 2 represents date + time
     * 1  represents date
     * 0  represents time
     * -1 unknown
     */
    private static int isReserved(short reserved) {
        /*
         * If it is the date of the Chinese type, convert its format to "yyyy-MM-dd".
         */
        Integer ck = reservedMap.get(reserved);
        return ck == null ? -1 : ck;
    }

    /*
    public static int isReserved(short reserved) {
        if (reserved == 22 || reserved == 188 || reserved == 189) {
            return 2;
        } else if (reserved >= 14&&reserved <= 17 || reserved >= 30 && reserved <= 31 ||
                reserved >= 57&&reserved <= 58 || reserved >= 180 && reserved <= 187 || reserved >= 190 && reserved <= 200)
            return 1;
        else if (reserved>=18&&reserved <=21 || reserved >= 32 && reserved <= 33 ||
                reserved >= 45&&reserved <= 47 || reserved >= 55 && reserved <= 56 ||
                reserved >= 201&&reserved <= 211) {
            return 0;
        } else {
            return -1;
        }
    }
    */

    /**
     * Determine whether it is Chinese date format
     *
     * @param isNotDate
     * @return
     */
    private static boolean isZnDateFormat(String isNotDate) {
        if (isNotDate == null || isNotDate.length() == 0) {
            return false;
        }
        return znMap.contains(isNotDate);
    }

    /**
     * The date judgement of POI is only applicable to the date format in Europe and America.
     * It does not support the Chinese date, and uses two ways(including isReserved and isDateFormat) to determine the date of the Chinese format.
     *
     * @param cell
     * @return
     */
    public static boolean isDateFormat(Cell cell) {
        if (cell == null) {
            return false;
        }
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
            return true;
        }
        if (isReserved(cell.getCellStyle().getDataFormat()) != -1) {
            return true;
        }
        return isZnDateFormat(cell.getCellStyle().getDataFormatString());
    }

    /**
     * handle some numeric cells, their style is as follows:
     * "^0(\.0+)?_\s$" matches the format like "0.0000_ "
     * "^0(\.0+)?_?[\s\)]?;(\[RED])?(\\\-|\\\()?0(\.0+)?(\\\s|\\\))?$" matches the format like "0.0000;[Red]0.0000", "0.0000_);[Red]\(0.0000\)", "0.0000_);\(0.0000\)" and "0.0000_ ;[Red]\-0.0000\ "
     * "^0(.0+)?%$" matches the format of the percentage, like "0.00%", "0%"
     * Notice: The BigDecimal is used here instead of the DecimalFormat. The DecimalFormat will lose precision.
     * @param cell
     * @return
     */
    private static String handleNumeric(Cell cell) {
        boolean isPercentage = false;
        String[] fItem = new String[0];
        String format = cell.getCellStyle().getDataFormatString();
        if(format == null) {
            return null;
        } else {
            format = format.toUpperCase();
        }
        if (format.matches("^0(\\.0+)?_\\s$")) {
            fItem = format.replaceAll("_\\s$", "").split("\\.");
        } else if (format.matches("^0(\\.0+)?_?[\\s)]?;(\\[RED])?(\\\\-|\\\\\\()?0(\\.0+)?(\\\\\\s|\\\\\\))?$")){
            fItem = format.replaceAll("_?[\\s)]?;(\\[RED])?(\\\\-|\\\\\\()?0(\\.0+)?(\\\\\\s|\\\\\\))?$", "").split("\\.");
        } else if (format.matches("^0(.0+)?%$")) {
            isPercentage = true;
            fItem = format.replaceAll("%$", "").split("\\.");
        }
        if(fItem.length > 0) {
            int precision = fItem.length == 2 ? fItem[1].length():0;
            BigDecimal bigDecimal = new BigDecimal(String.valueOf(cell.getNumericCellValue())).setScale(precision, BigDecimal.ROUND_HALF_UP);
            if(isPercentage) {
                percent.setMaximumFractionDigits(precision);
                return percent.format(bigDecimal.doubleValue());
            }else {
                return bigDecimal.stripTrailingZeros().toPlainString();
            }
        } else {
            return null;
        }
    }

    /**
     * return UUID whose length is 32
     * @return
     */
    public static String getUUID() {
        String uuid = UUID.randomUUID().toString().replace("-", "").toLowerCase();
        return uuid;
    }
}
