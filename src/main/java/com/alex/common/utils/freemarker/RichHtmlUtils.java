package com.alex.common.utils.freemarker;

import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class RichHtmlUtils {

    /**
     * 处理html返回html imagesBase64 imagesXmlHref
     *
     * @param html
     * @param docSrcLocationPrex file:///C:/8595226D
     * @param docSrcParent       file3405.files
     * @param nextPartId         01D214BC.6A592540
     * @return Map<String, String> handleHtml imagesBase64 imagesXmlHref
     */
    public static Map<String, String> computedImage(String html, String docSrcLocationPrex, String docSrcParent, String nextPartId) {
        Map<String, String> handleHtml = new HashMap<>();
        //处理后的html
        handleHtml.put("handleHtml", "");
        //base64
        handleHtml.put("imagesBase64", "");
        //ImagesXmlHref
        handleHtml.put("imagesXmlHref", "");
        if (StringUtils.isNotBlank(html)) {
            try {
                //				RichHtmlHandler handler = new RichHtmlHandler(html, "");
                RichHtmlHandler handler = new RichHtmlHandler(html, docSrcLocationPrex, docSrcParent, nextPartId);
                //				handler.setHtml(html);
                //				handler.setDocSrcLocationPrex(docSrcLocationPrex);
                //				handler.setDocSrcParent(docSrcParent);
                //				handler.setNextPartId(nextPartId);
                handler.setShapeidPrex("_x56fe__x7247__x0020");
                handler.setSpidPrex("_x0000_i");
                handler.setTypeid("#_x0000_t75");
                handler.handledHtml(false);
                String bodyBlock = handler.getHandledDocBodyBlock();
                //处理后的html
                handleHtml.put("handleHtml", bodyBlock);
                //base64
                handleHtml.put("imagesBase64", makeImagesBase64String(handler));
                //ImagesXmlHref
                handleHtml.put("imagesXmlHref", makeImagesXmlHrefString(handler));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return handleHtml;
    }


    /**
     * 生成图片ImagesBase64String
     *
     * @param handler
     * @return
     */
    private static String makeImagesBase64String(RichHtmlHandler handler) {
        String handledBase64Block = "";
        if (handler.getDocBase64BlockResults() != null
                && handler.getDocBase64BlockResults().size() > 0) {
            for (String item : handler.getDocBase64BlockResults()) {
                handledBase64Block += item + "\n";
            }
        }
        return handledBase64Block;
    }


    /**
     * 生成图片ImagesXmlHrefString
     *
     * @param handler
     * @return
     */
    private static String makeImagesXmlHrefString(RichHtmlHandler handler) {
        String xmlimaHref = "";
        if (handler.getXmlImgRefs() != null
                && handler.getXmlImgRefs().size() > 0) {
            for (String item : handler.getXmlImgRefs()) {
                xmlimaHref += item + "\n";
            }
        }
        return xmlimaHref;
    }
}
