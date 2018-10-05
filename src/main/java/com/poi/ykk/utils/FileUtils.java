package com.poi.ykk.utils;

import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;

import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;

public class FileUtils {

    public static final String UPLOAD_DIR = "upload";

    public static final String REPORT_DIR = "report";

    public static final String REPORT_TEMPLATE_DIR = "template";

    public static final SimpleDateFormat YMDateFormatter = new SimpleDateFormat("yyyy-MM");

    /**
     * 判断文件是否存在
     * @param path
     * @return
     */
    public static boolean isExist(Path path) {
        return path.toFile().exists();
    }

    public static boolean isExist(String fileName) {
        return new File(fileName).exists();
    }

    /**
     * 如果月份文件夹不存在则创建
     * @param fileName
     */
    public static void isExistOrCreate(String fileName) {
        Path path = Paths.get(UPLOAD_DIR, fileName);
        if (!isExist(path)) {
            path.toFile().mkdir();
        }
    }

    /**
     * 把当前时间的毫秒数转换成年月，作为文件夹名称
     * @param ms
     * @return
     */
    public static String transform(Long ms) {
        return YMDateFormatter.format(ms);
    }

    /**
     * 获取上传文件目录下资源的读取路径
     * @param fileName
     * @return
     */
    public static String getUploadFileAbsPath(String fileName) {
        if (fileName == null || fileName.length() == 0) {
            return "";
        }
        try {
            String[] _tokens = fileName.split("_");
            int _length = _tokens.length;
            Long timestamp = 0L;
            if (_length > 1) {
                String last_Token = _tokens[_length - 1];
                int indexOfDot = last_Token.indexOf('.');
                String timestampStr = last_Token;
                if (indexOfDot > -1) {
                    timestampStr = last_Token.substring(0, indexOfDot);
                }

                timestamp = Long.valueOf(timestampStr);
            }

            String dirStr = transform(timestamp);
            
            Path path = Paths.get(UPLOAD_DIR, dirStr, fileName);
            String fullFileName = path.toString();
            if (!isExist(path) && !dirStr.equals("")) { // 如果日期文件夹里没有，尝试从根路径查找
                fullFileName = Paths.get(UPLOAD_DIR, fileName).toString();
            }
            if (isExist(fullFileName)) {
                return fullFileName;
            }
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
        return fileName;
    }

    /**
     * 通用文件响应体
     * @param resourceLoader
     * @param filePath
     * @param fileName
     * @return
     */
    public static ResponseEntity buildExportFileResponse(
        ResourceLoader resourceLoader,
        Path filePath, 
        String fileName) {
        
        Resource resource;
        if (FileUtils.isExist(filePath)) {
            resource = resourceLoader.getResource("file:" + filePath.toString());
        } else {
            resource = resourceLoader.getResource("file:" + filePath.toString());
        }
        try {
            HttpHeaders headers = new HttpHeaders();
            headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
            headers.add("Pragma", "no-cache");
            headers.add("Expires", "0");
            // 保证导出的中文文件名正常
            // filename*=UTF-8''
            headers.add("Content-Disposition", 
                "attachment; filename*=UTF-8''" + URLEncoder.encode(fileName, "UTF-8"));
            return ResponseEntity
                .ok()
                .headers(headers)
                .contentLength(resource.contentLength())
                .contentType(MediaType.parseMediaType("application/force-download"))
                .body(resource);
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.notFound().build();
        }
    }
}
