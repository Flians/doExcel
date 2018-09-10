package com.unaware.poi.excel;

import com.unaware.poi.excel.exception.ReadException;
import com.unaware.poi.excel.util.MixedFile;
import org.apache.poi.hssf.record.RecordInputStream;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.List;

/**
 * EasyRefiner-step1
 *
 * @author mno
 * @date 2018/6/25 14:49
 */
public class UploadExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(UploadExcel.class);

    /**
     * 对文件进行清理、修剪
     *
     * @param file 文件
     * @return 文件信息
     */
    public List<MixedFile> analyzeFile(File file) {
        try {
            return analyzeExcelFile(file);
        } catch (NotOLE2FileException | RecordInputStream.LeftoverDataException e1) {
            // 特殊文件处理
            try {
                SpecialExcelConverter.toExcel(file);
                return analyzeExcelFile(file);
            } catch (NotOLE2FileException | RecordInputStream.LeftoverDataException e2) {
                LOGGER.error("#excel解析报错:{}", e2.getMessage());
                throw new ReadException("文件格式与文件扩展名的格式不一致，建议修改文件扩展名后重试");
            } catch (Exception e) {
                LOGGER.error("#excel解析报错:{}", e.getMessage());
                throw new ReadException("暂不支持您上传的文件类型_获取book失败");
            }
        } catch (Exception e) {
            LOGGER.error("#excel解析报错:{}", e.getMessage());
            throw new ReadException("暂不支持您上传的文件类型_获取book失败");
        }

    }

    /**
     * process Excel file and write its data into .csv file
     * use default parameters
     *
     * @param file 文件
     * @return 文件信息
     */
    public List<MixedFile> analyzeExcelFile(File file) throws Exception {
        try (SSConverter ssConverter = new SSConverter()) {
            ssConverter.enableAvailableInfo(true);
            ssConverter.path(file, -1, 10, -1);
            return ssConverter.getMixedFiles();
        }
    }
}
