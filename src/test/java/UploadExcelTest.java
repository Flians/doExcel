import com.unaware.poi.excel.UploadExcel;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.util.Arrays;
import java.util.Locale;

import static java.util.Objects.requireNonNull;

public class UploadExcelTest {
    @BeforeClass
    public static void init() {
        Locale.setDefault(Locale.CHINA);
    }

    private UploadExcel uploadExcel = new UploadExcel();


    /**
     * streaming method
     *
     */
    @Test
    public void testAnalyzeExcelFile() {
        File root = new File("src\\test\\resources\\testCase");
        Arrays.stream(requireNonNull(root.listFiles())).forEach(this::accept);
    }

    private void accept(File file) {
        if (!file.isDirectory()) {
            print_new(file);
        } else {
            Arrays.stream(requireNonNull(file.listFiles())).forEach(this::print_new);
        }
    }

    private void print_new(File item) {
        try {
            System.out.println("\n" + item.getName());
            Runtime r = Runtime.getRuntime();
            r.gc();
            long startMem = r.totalMemory() - r.freeMemory();
            long startTime = System.currentTimeMillis();
            System.out.println(uploadExcel.analyzeExcelFile(item));
            long endTime = System.currentTimeMillis();
            long orz = r.totalMemory() - r.freeMemory() - startMem;
            System.out.println("time used: " + (endTime - startTime) + ", memory used: " + orz);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}