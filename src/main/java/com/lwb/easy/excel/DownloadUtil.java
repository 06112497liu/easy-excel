package com.lwb.easy.excel;

import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Optional;
import java.util.function.Function;
import java.util.function.Predicate;

import static com.lwb.easy.excel.constant.Constant.*;
import static com.lwb.easy.excel.constant.Headers.CONTENT_DISPOSITION;
import static com.lwb.easy.excel.constant.Headers.USER_AGENT;
import static com.lwb.easy.excel.constant.MediaType.APPLICATION_OCTET_STREAM_VALUE;

/**
 * @author liuweibo
 * @date 2019/8/20
 */
public class DownloadUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(DownloadUtil.class);

    /**
     * 格式化下载文件名函数
     */
    private static final Function<String, String> FORMAT_FILE_NAME = (fileName) ->
        String.format("attachment; filename=\"%s.%s\"", fileName, XLSX);

    /**
     * 判断是否是ie内核浏览器断言
     */
    private static final Predicate<String> IS_IE = userAgent ->
        // 是否是ie浏览器
        userAgent.contains("MSIE")
            // 是否是edge浏览器
            || userAgent.contains("EDGE")
            // 是否是ie内核浏览器
            || userAgent.contains("TRIDENT");

    /**
     * 下载生成的临时文件
     */
    public static void download(HttpServletRequest request,
                                HttpServletResponse response,
                                String fileName,
                                InputStream inStream) throws Exception {

        // 设置下载文件名
        String newFileName =
            Optional.of(request.getHeader(USER_AGENT).toUpperCase())
                .filter(IS_IE)
                .map(t -> {
                    try {
                        return URLEncoder.encode(fileName, UTF_8);
                    } catch (UnsupportedEncodingException e) {
                        return EMPTY;
                    }
                }).orElse(new String(fileName.getBytes(UTF_8), ISO_8859_1));

        response.setContentType(APPLICATION_OCTET_STREAM_VALUE);
        response.setHeader(CONTENT_DISPOSITION, FORMAT_FILE_NAME.apply(newFileName));

        byte[] buffer = new byte[1024];
        try (OutputStream outStream = response.getOutputStream();
             PrintStream out = new PrintStream(outStream, true, UTF_8)) {
            int len;
            while ((len = inStream.read(buffer)) > 0) {
                out.write(buffer, 0, len);
                out.flush();
            }
        } finally {
            IOUtils.closeQuietly(inStream);
        }
    }
}
