package com.unaware.poi.excel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created by Administrator on 2018/7/9.
 */
public class ColorInfo {
    // color code in Excel
    private short color;
    //The alpha value of the color, which controls the transparency of the color
    private int A;
    //RGB: red
    private int R;
    //RGB: green
    private int G;
    //RGB: blue
    private int B;

    public int toRGB() {
        return this.R << 16 | this.G << 8 | this.B;
    }

    public java.awt.Color toAWTColor() {
        return new java.awt.Color(this.R, this.G, this.B, this.A);
    }

    public static ColorInfo fromARGB(int red, int green, int blue) {
        return new ColorInfo(0xff, red, green, blue);
    }

    public static ColorInfo fromARGB(int alpha, int red, int green, int blue) {
        return new ColorInfo(alpha, red, green, blue);
    }

    /**
     * constructor
     * @param a
     * @param r
     * @param g
     * @param b
     */
    public ColorInfo(int a, int r, int g, int b) {
        this.A = a;
        this.B = b;
        this.R = r;
        this.G = g;
    }

    /**
     * constructor
     * @param color
     * @param a
     * @param r
     * @param g
     * @param b
     */
    public ColorInfo(short color, int a, int r, int g, int b) {
        this.color = color;
        this.A = a;
        this.B = b;
        this.R = r;
        this.G = g;
    }

    /**
     * convert the color of excel(version including 97, 2003) to ColorInfo
     *
     * @param color
     * @return ColorInfo or null
     */
    public static ColorInfo excel97Color2UOF(Workbook book, short color) {
        if (book instanceof HSSFWorkbook) {
            HSSFWorkbook hb = (HSSFWorkbook) book;
            HSSFColor hc = hb.getCustomPalette().getColor(color);
            return excelColor2UOF(hc);
        } else if (book instanceof XSSFWorkbook) {
            XSSFWorkbook xw = (XSSFWorkbook) book;
            XSSFColor xc = xw.getTheme().getThemeColor(color);
            return excelColor2UOF(xc);
        }
        return null;
    }

    /**
     * convert the color of excel(version including 97, 2003 and 2007) to ColorInfo
     *
     * @param color
     * @return ColorInfo or null
     */
    public static ColorInfo excelColor2UOF(Color color) {
        if (color == null) {
            return null;
        }
        ColorInfo ci = null;
        if (color instanceof XSSFColor) {// .xlsx
            XSSFColor xc = (XSSFColor) color;
            byte[] rgb = xc.getRGB();
            if (rgb != null) {
                ci = ColorInfo.fromARGB(rgb[0], rgb[1], rgb[2]);
            }
        } else if (color instanceof HSSFColor) {// .xls
            HSSFColor hc = (HSSFColor) color;
            short[] s = hc.getTriplet();//RGB
            if (s != null) {
                ci = ColorInfo.fromARGB(s[0], s[1], s[2]);
            }
        }
        return ci;
    }

    /**
     * Each value is a signed binary number, which is converted as an unsigned number
     *
     * @param number
     * @return
     */
    private static int change(int number) {
        return ((number & 0x0f0) >> 4) * 16 + (number & 0x0f);
    }

    /**
     * convert org.apache.poi.ss.usermodel.Color to java.awt.Color
     * if Color is null, return java.awt.Color.white.
     *
     * @param col
     * @return java.awt.Color
     */
    public static java.awt.Color getColor(Color col) {
        if (col == null) {
            return java.awt.Color.white;
        }
        java.awt.Color color = null;
        if (col instanceof HSSFColor) {
            HSSFColor c = (HSSFColor) col;
            short[] triplet = c.getTriplet();
            if (c.getIndex() != 64)
                color = new java.awt.Color(triplet[0], triplet[1], triplet[2]);
        } else if (col instanceof XSSFColor) {
            XSSFColor c = (XSSFColor) col;
            byte[] rgbHex = c.getRGB();
            if (rgbHex != null)
                color = new java.awt.Color(change(rgbHex[0]), change(rgbHex[1]), change(rgbHex[2]));
        }
        if (color == null) {
            return java.awt.Color.white;
        }
        return color;
    }
}
