package com.alex.common.utils.freemarker;

import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.lang3.StringUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @Description:富文本Html处理器，主要处理图片及编码
 */
public class RichHtmlHandler {

    private static final Logger logger = LoggerFactory.getLogger(RichHtmlHandler.class);

    private Document doc = null;
    private String html;

    //	private String docSrcLocationPrex = "file:///C:/213792E5";   //在paper.ftl文件里面找到，检索“Content-Location”
    //	private String docSrcParent       = "airportWeeklyTemplate.files";     //在paper.ftl文件里面找到，检索“Content-Location”
    //	private String nextPartId         = "01D59300.D914CDC0";     //在paper.ftl文件里面找到，最末行
    private String docSrcLocationPrex;   //在paper.ftl文件里面找到，检索“Content-Location”
    private String docSrcParent;     //在paper.ftl文件里面找到，检索“Content-Location”
    private String nextPartId;     //在paper.ftl文件里面找到，最末行
    private String shapeidPrex = "_x56fe__x7247__x0020";
    private String spidPrex = "_x0000_i";
    private String typeid = "#_x0000_t75";

    private String handledDocBodyBlock;
    private List<String> docBase64BlockResults = new ArrayList<String>();
    private List<String> xmlImgRefs = new ArrayList<String>();

    private String srcPath = "";
    private String http = "http://";
    private String https = "https://";


    public RichHtmlHandler() {
    }


    public RichHtmlHandler(String html, String docSrcLocationPrex, String docSrcParent, String nextPartId) {
        this.html = html;
        this.docSrcLocationPrex = docSrcLocationPrex;
        this.docSrcParent = docSrcParent;
        this.nextPartId = nextPartId;
        doc = Jsoup.parse(wrappHtml(this.html));
    }


    public String getDocSrcLocationPrex() {
        return docSrcLocationPrex;
    }


    public void setDocSrcLocationPrex(String docSrcLocationPrex) {
        this.docSrcLocationPrex = docSrcLocationPrex;
    }


    public String getNextPartId() {
        return nextPartId;
    }


    public void setNextPartId(String nextPartId) {
        this.nextPartId = nextPartId;
    }


    public String getHandledDocBodyBlock() {
        String raw = WordHtmlGeneratorHelper.string2Ascii(doc.getElementsByTag("body").html());
        return raw.replace("=3D", "=").replace("=", "=3D");
    }


    public String getRawHandledDocBodyBlock() {
        String raw = doc.getElementsByTag("body").html();
        return raw.replace("=3D", "=").replace("=", "=3D");
    }


    public List<String> getDocBase64BlockResults() {
        return docBase64BlockResults;
    }


    public List<String> getXmlImgRefs() {
        return xmlImgRefs;
    }


    public String getShapeidPrex() {
        return shapeidPrex;
    }


    public void setShapeidPrex(String shapeidPrex) {
        this.shapeidPrex = shapeidPrex;
    }


    public String getSpidPrex() {
        return spidPrex;
    }


    public void setSpidPrex(String spidPrex) {
        this.spidPrex = spidPrex;
    }


    public String getTypeid() {
        return typeid;
    }


    public void setTypeid(String typeid) {
        this.typeid = typeid;
    }


    public String getDocSrcParent() {
        return docSrcParent;
    }


    public void setDocSrcParent(String docSrcParent) {
        this.docSrcParent = docSrcParent;
    }


    public Document getDoc() {
        return doc;
    }


    public void setDoc(Document doc) {
        this.doc = doc;
    }


    public void setHandledDocBodyBlock(String handledDocBodyBlock) {
        this.handledDocBodyBlock = handledDocBodyBlock;
    }


    public void setDocBase64BlockResults(List<String> docBase64BlockResults) {
        this.docBase64BlockResults = docBase64BlockResults;
    }


    public void setXmlImgRefs(List<String> xmlImgRefs) {
        this.xmlImgRefs = xmlImgRefs;
    }


    public String getSrcPath() {
        return srcPath;
    }


    public void setSrcPath(String srcPath) {
        this.srcPath = srcPath;
    }


    public String getHtml() {
        return html;
    }


    public void setHtml(String html) {
        this.html = html;
    }


    public RichHtmlHandler(String html, String srcPath) {
        this.html = html;
        this.srcPath = srcPath;
        doc = Jsoup.parse(wrappHtml(this.html));
        try {
            handledHtml(false);
        } catch (IOException e) {
            logger.debug("安全周报IOException： {}", e.getMessage());
        }
    }


    public void re_init(String html) {
        doc = null;
        doc = Jsoup.parse(wrappHtml(html));
        docBase64BlockResults.clear();
        xmlImgRefs.clear();
    }


    /**
     * @param @return
     * @return String
     * @throws IOException
     * @throws
     * @Description: 获得已经处理过的HTML文件
     */
    public void handledHtml(boolean isWebApplication)
            throws IOException {
        Elements imags = doc.getElementsByTag("img");
        if (imags == null || imags.size() == 0) {
            // 返回编码后字符串
            return;
            //handledDocBodyBlock = WordHtmlGeneratorHelper.string2Ascii(html);
        }

        // 转换成word mht 能识别图片标签内容，去替换html中的图片标签

        for (Element item : imags) {
            // 把文件取出来
            String src = item.attr("src");
            String srcRealPath = srcPath + src;

            //            String thepaths = RichHtmlHandler.class.getClassLoader().getResource("").toString();
            if (isWebApplication) {
                //                String contentPath=RequestResponseContext.getRequest().getContextPath();
                //                if(!StringUtils.isEmpty(contentPath)){
                //                    if(src.startsWith(contentPath)){
                //                        src=src.substring(contentPath.length());
                //                    }
                //                }
                //
                //                srcRealPath = RequestResponseContext.getRequest().getSession()
                //                        .getServletContext().getRealPath(src);

            }
            //判断图片是否是base64非网络资源
            String imageFielShortName = "";
            String fileTypeName = "";
            URL url = null;
            File imageFile = null;
            if (srcRealPath.contains("data:") || srcRealPath.contains("base64")) {
                imageFielShortName = DigestUtils.md5Hex(src + "weekly");
                //data:image/jpeg;base64,/9xasda.....
                String contentType = srcRealPath.substring(srcRealPath.indexOf(":") + 1, srcRealPath.indexOf(";"));
                fileTypeName = WordImageConvertor.getImageExtension(contentType);
            } else {
                //判断url是否可用
                if (StringUtils.isNotBlank(srcRealPath) && !srcRealPath.contains("http:") && !srcRealPath
                        .contains("https:") && !srcRealPath.contains("file:")) {
                    //Protocol + path
                    int path = srcRealPath.indexOf("//");
                    int startIndex = 0;
                    if (path != -1) {
                        startIndex = path + 2;
                    }
                    String srcRealPathHttp = http + srcRealPath.substring(startIndex);
                    String srcRealPathHttps = https + srcRealPath.substring(startIndex);
                    url = urlWithTimeOut(srcRealPath, srcRealPathHttp, srcRealPathHttps, 5000);
                    if (url != null) {
                        srcRealPath = url.getProtocol() + "://" + url.getHost() + url.getFile();
                        imageFile = new File(url.getPath());
                        imageFielShortName = imageFile.getName();
                        fileTypeName = WordImageConvertor.getFileSuffix(url.getPath());
                    }
                } else {
                    imageFile = new File(srcRealPath);
                    imageFielShortName = imageFile.getName();
                    fileTypeName = WordImageConvertor.getFileSuffix(srcRealPath);
                }
            }

            String docFileName = "image" + UUID.randomUUID().toString() + "." + fileTypeName;
            String srcLocationShortName = docSrcParent + "/" + docFileName;

            String styleAttr = item.attr("style"); // 样式
            //高度
            String imagHeightStr = item.attr("height");
            if (StringUtils.isEmpty(imagHeightStr)) {
                imagHeightStr = getStyleAttrValue(styleAttr, "height");
            }
            //宽度
            String imagWidthStr = item.attr("width");

            if (StringUtils.isEmpty(imagWidthStr)) {
                imagWidthStr = getStyleAttrValue(styleAttr, "width");
            }

            imagHeightStr = imagHeightStr.replace("px", "");
            imagWidthStr = imagWidthStr.replace("px", "");
            if (StringUtils.isEmpty(imagHeightStr)) {
                //去得到默认的文件高度
                imagHeightStr = "0";
            }
            if (imagHeightStr.equals("auto")) {
                //去得到默认的文件高度
                imagHeightStr = "0";
            }
            if (StringUtils.isEmpty(imagWidthStr)) {
                imagWidthStr = "0";
            }
            if (imagWidthStr.equals("auto")) {
                imagWidthStr = "0";
            }
            float imageHeight = Float.parseFloat(imagHeightStr);
            float imageWidth = Float.parseFloat(imagWidthStr);
            //比例缩放
            if (imageHeight > 234) {
                imageHeight = 234.0f;
            }
            if (imageWidth > 234) {
                imageWidth = 234.0f;
            }
            // 得到文件的word mht的body块
            String handledDocBodyBlock = WordImageConvertor.toDocBodyBlock(srcRealPath,
                    imageFielShortName, imageHeight, imageWidth, styleAttr,
                    srcLocationShortName, shapeidPrex, spidPrex, typeid);

            //这里的顺序有点问题：应该是替换item，而不是整个后面追加
            //doc.rreplaceAll(item.toString(), handledDocBodyBlock);
            item.after(handledDocBodyBlock);
            //            item.parent().append(handledDocBodyBlock);
            item.remove();
            // 去替换原生的html中的img
            String base64Content = "";
            if (srcRealPath.contains("data:") || srcRealPath.contains("base64")) {
                base64Content = srcRealPath.substring(srcRealPath.indexOf(";") + 1);
            } else {
                base64Content = WordImageConvertor.imageToBase64(srcRealPath);
            }
            String contextLoacation = docSrcLocationPrex + "/" + docSrcParent + "/" + docFileName;

            String docBase64BlockResult = WordImageConvertor.generateImageBase64Block(nextPartId, contextLoacation,
                    fileTypeName, base64Content);
            docBase64BlockResults.add(docBase64BlockResult);

            String imagXMLHref = "<o:File HRef=3D\"" + docFileName + "\"/>";
            xmlImgRefs.add(imagXMLHref);

        }

    }


    private String getStyleAttrValue(String style, String attributeKey) {
        if (StringUtils.isEmpty(style)) {
            return "";
        }

        // 以";"分割
        String[] styleAttrValues = style.split(";");
        for (String item : styleAttrValues) {
            // 在以 ":"分割
            String[] keyValuePairs = item.split(":");
            if (attributeKey.equals(keyValuePairs[0])) {
                return keyValuePairs[1];
            }
        }

        return "";
    }


    private String wrappHtml(String html) {
        // 因为传递过来都是不完整的doc
        StringBuilder sb = new StringBuilder();
        sb.append("<html>");
        sb.append("<body>");
        sb.append(html);

        sb.append("</body>");
        sb.append("</html>");
        return sb.toString();
    }


    public String getData(List<String> list) {
        String data = "";
        if (list != null && list.size() > 0) {
            for (String string : list) {
                data += string + "\n";
            }
        }
        return data;
    }


    /**
     * 网络地址是否可用
     *
     * @param timeOutMillSeconds
     */
    public static URL urlWithTimeOut(String srcUrl, String httpUrl, String httpsUrl, int timeOutMillSeconds) {
        URL url = null;
        HttpURLConnection co = null;
        try {
            url = new URL(srcUrl);
            co = (HttpURLConnection) url.openConnection();
            co.setRequestMethod("GET");
            co.setConnectTimeout(timeOutMillSeconds);
            co.connect();
            logger.info("srcUrl 连接可用");
        } catch (Exception e1) {
            logger.debug("srcUrl 连接打不开!,{}", e1.getMessage());
            try {
                url = new URL(httpsUrl);
                co = (HttpURLConnection) url.openConnection();
                co.setConnectTimeout(timeOutMillSeconds);
                co.connect();
                logger.info("httpsUrl 连接可用");
            } catch (Exception e2) {
                logger.debug("httpsUrl 连接打不开!,{}", e2.getMessage());
                try {
                    url = new URL(httpUrl);
                    co = (HttpURLConnection) url.openConnection();
                    co.setConnectTimeout(timeOutMillSeconds);
                    co.connect();
                    logger.info("httpUrl 连接可用");
                } catch (Exception e3) {
                    logger.debug("httpUrl 连接打不开!,{}", e3.getMessage());
                }
            }
        }
        return url;
    }
}