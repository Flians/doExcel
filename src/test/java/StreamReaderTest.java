import com.unaware.poi.excel.streamreader.StreamReader;
import com.unaware.poi.excel.util.DataUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.BeforeClass;
import org.junit.Test;

import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

public class StreamReaderTest {
    @BeforeClass
    public static void init() {
        Locale.setDefault(Locale.CHINA);
    }
    @Test
    public void testLambda() {
        List<Point> test = new ArrayList<Point>(){
            {
                add(new Point(1,2));
                add(new Point(1,2));
                add(null);
                add(new Point(1,2));
            }
        };
        test.forEach(point -> System.out.println(point.x));
    }

    @Test
    public void testEmptyCellShouldHaveGeneralStyle() {
        File root = new File("src\\test\\resources\\testCase");
        Arrays.asList(Objects.requireNonNull(root.listFiles())).forEach(this::accept);
    }

    private void accept(File file) {
        if (!file.isDirectory()) {
            try {
                printExcel(file);
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {
            Arrays.asList(Objects.requireNonNull(file.listFiles())).forEach(item -> {
                try {
                    System.out.println("\n" + item.getName());
                    printExcel(item);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
        }
    }

    private void printExcel(File file) throws Exception {
        try (
                /*
                  print all information to file
                 */
                FileOutputStream fos = new FileOutputStream(new File("src\\test\\resources\\output\\" + file.getName().substring(0, file.getName().lastIndexOf('.')) + ".csv"));
                BufferedWriter fileWriter = new BufferedWriter(new OutputStreamWriter(fos));
                /*
                  read all information from excel file
                 */
                InputStream is = new FileInputStream(file);
                //Workbook workbook = StreamReader.builder().sstCacheSize(10).rowCacheSize(10).bufferSize(1024).sheetIndex(-1).open(is, StreamReader.Builder.ExcelType.XLSX)
                Workbook workbook = StreamReader.builder().sstCacheSize(10).rowCacheSize(20).sheetIndex(-1).open(file)
        ) {
            // write bom in header to solve the problem of Chinese chaotic code in Excel
            byte[] uft8bom={(byte)0xef,(byte)0xbb,(byte)0xbf};
            fos.write(uft8bom);

            System.out.println(file.getName() + " has Number of the sheets: " + workbook.getNumberOfSheets());
            fileWriter.write("cell_index:,border_top_type:,border_top_color:,border_bottom_type:,border_bottom_color:,border_left_type:,border_left_color:,border_right_type:,border_right_color:," +
                    "background_color:," + "font_type:,font_color:,font_size:,font_bold:," + "cell_type:,cell_value_unformat:,cell_value_format:\n");
            workbook.forEach(sheet -> {
//                System.out.println(">>>>>>>>>>> " + sheet.getSheetName());
                sheet.forEach(row -> row.forEach(c -> {
                    try {
                        //cell index
                        fileWriter.write(c.getAddress().toString() + ",");
                        //border line
                        fileWriter.write(c.getCellStyle().getBorderTopEnum() + "," + c.getCellStyle().getTopBorderColor() + "," +
                                c.getCellStyle().getBorderBottomEnum() + "," + c.getCellStyle().getBottomBorderColor() + "," +
                                c.getCellStyle().getBorderLeftEnum() + "," + c.getCellStyle().getLeftBorderColor() + "," +
                                c.getCellStyle().getBorderRightEnum() + "," + c.getCellStyle().getRightBorderColor() + ",");
                        //background color
                        if(c.getCellStyle().getFillForegroundColorColor() instanceof XSSFColor) {
                            if (c.getCellStyle().getFillForegroundColorColor() != null) {
                                XSSFColor xc = (XSSFColor) c.getCellStyle().getFillForegroundColorColor();
                                fileWriter.write(xc.getARGBHex() + ",");
                            } else {
                                fileWriter.write(c.getCellStyle().getFillForegroundColor() + ",");
                            }
                        }else {
                            if (c.getCellStyle().getFillForegroundColorColor() != null) {
                                HSSFColor xc = (HSSFColor) c.getCellStyle().getFillForegroundColorColor();
                                fileWriter.write(xc.getHexString() + ",");
                            }else {
                                fileWriter.write(c.getCellStyle().getFillForegroundColor() + ",");
                            }
                        }

                        //font style
                        short index = c.getCellStyle().getFontIndex();
                        //font name
                        fileWriter.write(workbook.getFontAt(index).getFontName() + " " + index + ",");
                        //font color
                        Font font = workbook.getFontAt(index);
                        if(font instanceof XSSFFont){
                            XSSFFont xf = (XSSFFont)font;
                            if(xf.getXSSFColor() != null) {
                                fileWriter.write(xf.getXSSFColor().getARGBHex() + ",");
                            }else {
                                fileWriter.write(font.getColor() + ",");
                            }
                        }else {
                            HSSFFont hf = (HSSFFont)font;
                            if(hf.getHSSFColor((HSSFWorkbook) workbook) != null) {
                                fileWriter.write(hf.getHSSFColor((HSSFWorkbook) workbook).getHexString() + ",");
                            }else {
                                fileWriter.write(font.getColor() + ",");
                            }
                        }

                        //font size
                        fileWriter.write(workbook.getFontAt(index).getFontHeightInPoints() + ",");
                        //font bold
                        if (workbook.getFontAt(index).getBold())
                            fileWriter.write("Y,");
                        else
                            fileWriter.write("N,");
                        //cell data format
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        String value;
                        switch (c.getCellTypeEnum()) {
                            case FORMULA:
                                value = "FORMULA," + c.getCellFormula();
                                break;

                            case NUMERIC:
                                if(DataUtil.isDateFormat(c)){
                                    value = "DATA," + sdf.format(c.getDateCellValue()) + " " + HSSFDateUtil.isCellDateFormatted(c);
                                } else {
                                    value = "NUMERIC," + c.getNumericCellValue();
                                }
                                break;

                            case STRING:
                                value = "STRING," + c.getStringCellValue().replaceAll("[\\r\\n,]", " ");
                                break;

                            case BLANK:
                                value = "BLANK,";
                                break;

                            case BOOLEAN:
                                value = "BOOLEAN," + c.getBooleanCellValue();
                                break;

                            case ERROR:
                                value = "ERROR," + c.getErrorCellValue();
                                break;

                            default:
                                value = "UNKNOWN," + c.getCellType();
                        }
                        fileWriter.write(value + "," + DataUtil.getCellValue(c).replaceAll("[\\r\\n,]", " ") + "," + c.getCellStyle().getDataFormat() + "," + c.getCellStyle().getDataFormatString() + "\n");
                    } catch (IOException e){
                        e.printStackTrace();
                    }
                }));
                try {
                    fileWriter.write(sheet.getSheetName() + " Sheet has " + sheet.getPhysicalNumberOfRows() + " lines\n");
                    //merged cells
                    fileWriter.write(sheet.getNumMergedRegions() + " --> " + sheet.getMergedRegions().toString() + "\n");
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            fileWriter.flush();
        }
    }
}
