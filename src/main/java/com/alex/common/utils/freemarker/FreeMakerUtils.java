package com.alex.common.utils.freemarker;

import com.alex.common.constant.FtlConstants;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import lombok.extern.slf4j.Slf4j;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

@Slf4j
public class FreeMakerUtils {
    public static final String CHARSET = "UTF-8";

    /**
     * 生成word文件
     *
     * @param dataMap  生成doc参数
     * @param filePath 生成doc文件地址
     * @param ftlName  ftl模板名称
     */
    public static void createDoc(Map<String, Object> dataMap, String filePath, String ftlName) {
        //输出文档路径及名称
        File outFile = new File(filePath);
        if (!outFile.getParentFile().exists()) {
            outFile.getParentFile().mkdirs();
        }
        try (
                FileOutputStream fos = new FileOutputStream(outFile);
                OutputStreamWriter writer = new OutputStreamWriter(fos, CHARSET)
        ) {
            Configuration configuration = new Configuration(Configuration.VERSION_2_3_28);
            configuration.setDefaultEncoding(CHARSET);
            configuration.setLocale(Locale.CHINESE);
            //加载classpath路径下的ftl模板
            configuration.setClassForTemplateLoading(FreeMakerUtils.class, FtlConstants.FTL_DIR);
            Template t = configuration.getTemplate(ftlName, CHARSET);
            t.process(dataMap, writer);
        } catch (FileNotFoundException e) {
            log.info("ftl模板不存在,生成doc失败,：{}", e.getMessage());
        } catch (IOException e) {
            log.info("IO异常,生成doc失败：{}", e.getMessage());
        } catch (TemplateException e) {
            log.info("ftl模板异常：{}", e.getMessage());
        } catch (Exception e) {
            log.info("生成doc失败：{}", e.getMessage());
        }
    }

    /**
     * 生成图片信息
     * base64 ext width height
     *
     * @param imgFile 文件绝对路径
     * @param type    文件ext
     * @return
     */
    public static Map<String, String> getImageStr(String imgFile, String type) {
        Map<String, String> imageInfo = new HashMap<>();
        if (imgFile == null || imgFile.trim() == "") {
            return imageInfo;
        }
        InputStream in = null;
        InputStream im = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            log.error("getImageInfo I/O exception " + e.getMessage(), e);
        } finally {
            if (in != null || im != null) {
                try {
                    in.close(); // 关闭流
                } catch (IOException e) {
                    log.error("getImageInfo I/O exception " + e.getMessage(), e);
                }
            }
        }
        BASE64Encoder encoder = new BASE64Encoder();
        imageInfo.put("base64", encoder.encode(data));
        imageInfo.put("type", type);
        return imageInfo;
    }

    public static Map<String, String> getImageStrByNetwork(String imgFile, String type) {
        String base64 = "";
        ByteArrayOutputStream outPut = new ByteArrayOutputStream();
        byte[] data = new byte[1024];
        try {
            // 创建URL
            URL url = new URL(imgFile);
            // 创建链接
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(10 * 1000);
            if (conn.getResponseCode() != HttpURLConnection.HTTP_OK) {
                base64 = "";//连接失败/链接失效/图片不存在
            }
            InputStream inStream = conn.getInputStream();
            int len = -1;
            while ((len = inStream.read(data)) != -1) {
                outPut.write(data, 0, len);
            }
            inStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 对字节数组Base64编码
        BASE64Encoder encoder = new BASE64Encoder();
        base64 = encoder.encode(outPut.toByteArray());
        Map<String, String> imageInfo = new HashMap<>();
        imageInfo.put("base64", base64);
        imageInfo.put("type", type);
        return imageInfo;
    }
}
