package com.lwb.easy.excel;

import com.lwb.easy.excel.annotation.Export;
import com.lwb.easy.excel.exception.ExcelException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Method;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.UUID;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

import static com.lwb.easy.excel.constant.Constant.*;

/**
 * excel工具类
 * </p>
 * 根据提供的数据和配置，生成临时excel临时文件供后续下载
 * @author liuweibo
 * @date 2019/8/20
 */
public class ExcelUtil {

    /**
     * 临时excel文件存放位置
     */
    private static final String TEMP_EXCEL_PATH = "temp";
    private static final String CLASSPATH_URL_PREFIX = "classpath:";
    private static Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    private static ThreadPoolExecutor EXECUTOR;

    static {
        int coreSize = Runtime.getRuntime().availableProcessors();
        EXECUTOR = new ThreadPoolExecutor(
            coreSize,
            coreSize << 1,
            200,
            TimeUnit.SECONDS,
            new LinkedBlockingQueue<>(coreSize << 2),
            new ThreadPoolExecutor.CallerRunsPolicy()
        );
    }

    /**
     * 生成excel文件，并保存为临时文件，供后续下载
     * @param data 数据
     * @return 文件名
     */
    public static String save(List<?> data) {
        Method method = ExcelHelper.getMethod(Export.class, Thread.currentThread().getStackTrace());
        ExcelConfig config = ExcelHelper.parseYml(method);
        // 配置完整性校验
        config.validate();
        return save(generateExcel(config, data), config);
    }

    /**
     * 直接通过文件名下载生成的excel文件
     * @param fileName 文件名
     * @param request  请求
     * @param response 响应
     */
    public static void download(String fileName, HttpServletRequest request, HttpServletResponse response) {
        String fileFullName = getFileFullPath(fileName);

        try {
            FileInputStream inputStream = new FileInputStream(fileFullName);
            DownloadUtil.download(
                request,
                response,
                fileName.split(POINT)[0],
                inputStream
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelException(e.getMessage());
        } finally {
            // 异步删除文件
            EXECUTOR.execute(() ->
                Optional.of(new File(fileFullName))
                    // 文件是否存在
                    .filter(File::exists)
                    // 删除文件
                    .filter(File::delete)
                    .ifPresent(file -> LOGGER.debug(String.format("file %s deleted!", fileFullName)))
            );
        }
    }

    /**
     * 生成excel并下载
     * @param data     数据
     * @param request  请求
     * @param response 响应
     */
    public static void download(List<?> data, HttpServletRequest request, HttpServletResponse response) {
        download(data, ExcelHelper.parseConfig(), request, response);
    }

    /**
     * 生成excel并下载
     * @param data     数据
     * @param config   excel配置文件
     * @param request  请求
     * @param response 响应
     */
    private static void download(List<?> data, ExcelConfig config, HttpServletRequest request, HttpServletResponse response) {
        SXSSFWorkbook book = generateExcel(config, data);
        // 输出文件流到response
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            book.write(out);
            DownloadUtil.download(
                request,
                response,
                config.getFileName(),
                new ByteArrayInputStream(out.toByteArray())
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelException(e.getMessage());
        }
    }

    /**
     * 生成excel
     * @param config excel配置
     * @param data   数据
     * @return 文件名
     */
    private static SXSSFWorkbook generateExcel(ExcelConfig config, List<?> data) {
        SXSSFWorkbook book = new SXSSFWorkbook();
        SXSSFSheet sheet = book.createSheet();
        // 表头样式
        CellStyle headerStyle = ExcelStyle.headerStyle(book);
        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();
        // 绘制表头
        config.getHeaders()
            .forEach(headers -> {
                // 获取有记录的行数（最后有数据的行是第n行，前面有m行是空行没数据，则返回n-m）
                SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                headers.forEach(header -> {
                    String name = header.getName();
                    String merge = header.getMergeIndex();
                    Cell cell = row.createCell(row.getPhysicalNumberOfCells());

                    // 是否合并单元格
                    if (merge != null) {
                        String[] index = merge.split(",");
                        CellRangeAddress cellAddresses = new CellRangeAddress(
                            Integer.parseInt(index[0]),
                            Integer.parseInt(index[1]),
                            Integer.parseInt(index[2]),
                            Integer.parseInt(index[3])
                        );
                        sheet.addMergedRegion(cellAddresses);
                        // 收集合并的单元格，用于后续设置合并后的样式，防止合并后单元格样式丢失
                        cellRangeAddresses.add(cellAddresses);
                    }
                    cell.setCellValue(name);
                    cell.setCellStyle(headerStyle);
                });
            });

        // excel设置单元格值
        Optional.ofNullable(data)
            .filter(CollectionUtils::isNotEmpty)
            .ifPresent(list -> {
                SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                list.forEach(item -> config.getFields()
                    .forEach(fieldName -> {
                        SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
                        try {
                            cell.setCellValue(ExcelHelper.getFieldValue(item, fieldName));
                        } catch (Exception e) {
                            LOGGER.error(e.getMessage(), e);
                            throw new ExcelException(e.getMessage());
                        }
                    }));
            });

        // 设置合并单元格后的单元格样式
        ExcelStyle.setCellRangeAddress(cellRangeAddresses, sheet);

        // 冻结表头
        Optional.ofNullable(config.getFreezePaneIndex())
            .filter(StringUtils::isNotEmpty)
            .filter(s -> s.contains(COMMA))
            .map(s -> {
                String[] index = s.split(COMMA);
                sheet.createFreezePane(
                    Integer.parseInt(index[0]),
                    Integer.parseInt(index[1]),
                    Integer.parseInt(index[2]),
                    Integer.parseInt(index[3])
                );
                return EMPTY;
            })
            // 默认冻结表头行数
            .orElseGet(() -> {
                sheet.createFreezePane(0, config.getHeaders().size());
                return EMPTY;
            });
        return book;
    }

    /**
     * 生成临时文件，供后续下载
     * @param book   excel文件
     * @param config excel配置
     */
    private static String save(Workbook book, ExcelConfig config) {
        // 生成唯一文件名
        String fileName = String.format("%s_%s.%s", config.getFileName(), UUID.randomUUID(), XLSX);

        FileOutputStream out = null;
        try {
            String fileFullPath = getFileFullPath(fileName);
            File file = new File(fileFullPath);
            // 创建临时文件夹
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdir();
            }
            // 创建临时文件
            if (!file.exists()) {
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            book.write(out);
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelException(e.getMessage());
        } finally {
            IOUtils.closeQuietly(out);
        }
        return fileName;
    }

    /**
     * 获得临时文件全路径
     * @param fileName 临时文件名包含后缀
     * @return 全路径文件名
     */
    public static String getFileFullPath(String fileName) {
        return getClassPathURL() + File.separator + TEMP_EXCEL_PATH + File.separator + fileName;
    }

    /**
     * 获取classpath路径
     * @return
     */
    private static String getClassPathURL() {
        String path = CLASSPATH_URL_PREFIX.substring(CLASSPATH_URL_PREFIX.length());
        ClassLoader cl = getDefaultClassLoader();
        URL url = (cl != null ? cl.getResource(path) : ClassLoader.getSystemResource(path));
        if (url == null) {
            String description = "class path resource [" + path + "]";
            throw new ExcelException(description +
                " cannot be resolved to URL because it does not exist");
        }
        return url.getPath();
    }

    /**
     * 获取默认的类加载器
     * @return 类加载器
     */
    private static ClassLoader getDefaultClassLoader() {
        ClassLoader cl = null;
        try {
            cl = Thread.currentThread().getContextClassLoader();
        } catch (Throwable ex) {
        }
        if (cl == null) {
            cl = ExcelHelper.class.getClassLoader();
            if (cl == null) {
                try {
                    cl = ClassLoader.getSystemClassLoader();
                } catch (Throwable ex) {
                }
            }
        }
        return cl;
    }
}
