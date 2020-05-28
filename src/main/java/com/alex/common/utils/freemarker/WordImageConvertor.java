package com.alex.common.utils.freemarker;

import org.apache.commons.codec.binary.Base64;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import sun.misc.BASE64Decoder;
import sun.misc.BASE64Encoder;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.UUID;

/**
 * @Description:WORD 文档图片转换器
 */
public class WordImageConvertor {

	//private static Const WORD_IMAGE_SHAPE_TYPE_ID="";

	private static final Logger logger = LoggerFactory.getLogger(WordImageConvertor.class);



	/**
	 * @param @param  imageSrc 文件路径
	 * @param @return
	 * @return String
	 * @throws IOException
	 * @throws
	 * @Description: 将图片转换成base64编码的字符串
	 */
	public static String imageToBase64(String imageSrc ) throws IOException {
		//判断文件是否存在
		//		File file = new File(imageSrc);
		//		if (!file.exists()) {
		//			throw new FileNotFoundException("文件不存在！");
		//		}
		StringBuilder pictureBuffer = new StringBuilder();
		try {
			URL url = new URL(imageSrc);
			HttpURLConnection conn = (HttpURLConnection) url.openConnection();
			conn.setRequestMethod("GET");
			conn.setConnectTimeout(5 * 1000);
			try (InputStream input = conn.getInputStream()) {
				// 通过输入流获取图片数据
				//				InputStream input = conn.getInputStream();
				//		FileInputStream input = new FileInputStream(file);
				ByteArrayOutputStream out = new ByteArrayOutputStream();

				//读取文件

				//BufferedInputStream bi=new BufferedInputStream(in);
				Base64 base64 = new Base64();
				BASE64Encoder encoder = new BASE64Encoder();
				byte[] temp = new byte[1024];
				for (int len = input.read(temp); len != -1; len = input.read(temp)) {
					out.write(temp, 0, len);
					//out(pictureBuffer.toString());
					//out.reset();
				}
				pictureBuffer.append(new String(base64.encodeBase64Chunked(out.toByteArray())));
				//pictureBuffer.append(encoder.encodeBuffer(out.toByteArray()));


        /*byte[] data=new byte[input.available()];
        input.read(data);
        pictureBuffer.append(base64.encodeBase64String (data));*/

				input.close();
        /*BASE64Decoder decoder=new BASE64Decoder();
        FileOutputStream write = new FileOutputStream(new File("c:\\test2.jpg"));
        //byte[] decoderBytes = decoder.decodeBuffer (pictureBuffer.toString());
        byte[] decoderBytes = base64.decodeBase64(pictureBuffer.toString());
        write.write(decoderBytes);
        write.close();*/

				return pictureBuffer.toString();
			} catch (Exception e) {
				logger.debug("安全周报 imageToBase64 Exception： {}", e.getMessage());
			}
		} catch (Exception e) {
			logger.debug("安全周报 imageToBase64 Exception： {}", e.getMessage());
		}
		return pictureBuffer.toString();
	}



	public static String toDocBodyBlock(
            String imageFilePath,
            String imageFielShortName,
            float imageHeight,
            float imageWidth,
            String imageStyle,
            String srcLocationShortName,
            String shapeidPrex, String spidPrex, String typeid ) {
		//shapeid
		//mht文件中针对shapeid的生成好像规律，其内置的生成函数没法得知，但是只要保证其唯一就行
		//这里用前置加32位的uuid来保证其唯一性。
		String shapeid = shapeidPrex;
		shapeid += UUID.randomUUID().toString();

		//spid ,同shapeid处理
		String spid = spidPrex;
		spid += UUID.randomUUID().toString();


    /*    <!--[if gte vml 1]><v:shape id=3D"_x56fe__x7247__x0020_0" o:spid=3D"_x0000_i10=
                26"
                   type=3D"#_x0000_t75" alt=3D"725017921264249223.jpg" style=3D'width:456.7=
                5pt;
                   height:340.5pt;visibility:visible;mso-wrap-style:square'>
                   <v:imagedata src=3D"file9462.files/image001.jpg" o:title=3D"725017921264=
                249223"/>
                  </v:shape><![endif]--><![if !vml]><img width=3D609 height=3D454
                  src=3D"file9462.files/image002.jpg" alt=3D725017921264249223.jpg v:shapes=
                =3D"_x56fe__x7247__x0020_0"><![endif]>*/
		StringBuilder sb1 = new StringBuilder();

		sb1.append(" <!--[if gte vml 1]>");
		sb1.append("<v:shape id=3D\"" + shapeid + "\"");
		sb1.append("\n");
		sb1.append(" o:spid=3D\"" + spid + "\"");
		sb1.append(" type=3D\"" + typeid + "\" alt=3D\"" + imageFielShortName + "\"");
		sb1.append("\n");

		if (imageFilePath.contains("data:") || imageFilePath.contains("base64")) {
			sb1.append(
					" style=3D' " + generateImageBodyBlockStyleAttrByBase64(imageFilePath, imageHeight, imageWidth)
							+ imageStyle
							+ "'");
		} else {
			sb1.append(
					" style=3D' " + generateImageBodyBlockStyleAttr(imageFilePath, imageHeight, imageWidth) + imageStyle
							+ "'");
		}

		sb1.append(">");
		sb1.append("\n");
		sb1.append(" <v:imagedata src=3D\"" + srcLocationShortName + "\"");
		sb1.append("\n");
		sb1.append(" o:title=3D\"" + imageFielShortName.split("\\.")[0] + "\"");
		sb1.append("/>");
		sb1.append("</v:shape>");
		sb1.append("<![endif]-->");

		//以下是为了兼容游览器显示时的效果，但是如果是纯word阅读的话没必要这么做。
    /*    StringBuilder sb2=new StringBuilder();
        sb2.append(" <![if !vml]>");

        sb2.append("<img width=3D"+imageWidth +" height=3D" +imageHeight +
                  " src=3D\"" + srcLocationShortName +"\" alt=" +imageFielShortName+
                  " v:shapes=3D\"" + shapeid +"\">");

        sb2.append("<![endif]>");*/

		//return sb1.toString()+sb2.toString();
		return sb1.toString();
	}



	/**
	 * @param @param  nextPartId
	 * @param @param  contextLoacation
	 * @param @param  ContentType
	 * @param @param  base64Content
	 * @param @return
	 * @return String
	 * @throws
	 * @Description: 生成图片的base4块
	 */
	public static String generateImageBase64Block(String nextPartId, String contextLoacation,
                                                  String fileTypeName, String base64Content) {
        /*--=_NextPart_01D188DB.E436D870
                Content-Location: file:///C:/70ED9946/file9462.files/image001.jpg
                Content-Transfer-Encoding: base64
                Content-Type: image/jpeg

                base64Content
        */

		StringBuilder sb = new StringBuilder();
		sb.append("\n");
		sb.append("\n");
		sb.append("------=_NextPart_" + nextPartId);
		sb.append("\n");
		sb.append("Content-Location: " + contextLoacation);
		sb.append("\n");
		sb.append("Content-Transfer-Encoding: base64");
		sb.append("\n");
		sb.append("Content-Type: " + getImageContentType(fileTypeName));
		sb.append("\n");
		sb.append("\n");
		sb.append(base64Content);

		return sb.toString();
	}



	private static String generateImageBodyBlockStyleAttr(String imageFilePath, float height, float width) {
		StringBuilder sb = new StringBuilder();
		try {
			URL url = new URL(imageFilePath);
			HttpURLConnection conn = (HttpURLConnection) url.openConnection();
			conn.setRequestMethod("GET");
			conn.setConnectTimeout(5 * 1000);
			try (InputStream input = conn.getInputStream();) {
				sb = getImageInfo(input, height, width);

				//				sourceImg = ImageIO.read(input);
				//				if (height == 0) {
				//					height = sourceImg.getHeight();
				//				}
				//				if (width == 0) {
				//					width = sourceImg.getWidth();
				//				}
				//				//将像素转化成pt
				//				BigDecimal heightValue = new BigDecimal(height * 12 / 16);
				//				heightValue = heightValue.setScale(2, BigDecimal.ROUND_HALF_UP);
				//				BigDecimal widthValue = new BigDecimal(width * 12 / 16);
				//				widthValue = widthValue.setScale(2, BigDecimal.ROUND_HALF_UP);
				//
				//				sb.append("height:" + heightValue + "pt;");
				//				sb.append("width:" + widthValue + "pt;");
				//				sb.append("visibility:visible;");
				//				sb.append("mso-wrap-style:square; ");
			} catch (FileNotFoundException e) {
				logger.debug("安全周报图片错误FileNotFoundException： {}", e.getMessage());
			} catch (IOException e) {
				logger.debug("安全周报图片错误IOException： {}", e.getMessage());
			}
		} catch (Exception e) {
			logger.debug("安全周报 generateImageBodyBlockStyleAttr Exception： {}", e.getMessage());
		}

		return sb.toString();
	}



	private static String getImageContentType(String fileTypeName) {
		String result = "image/jpeg";
		//http://tools.jb51.net/table/http_content_type
		if (fileTypeName.equals("tif") || fileTypeName.equals("tiff")) {
			result = "image/tiff";
		} else if (fileTypeName.equals("fax")) {
			result = "image/fax";
		} else if (fileTypeName.equals("gif")) {
			result = "image/gif";
		} else if (fileTypeName.equals("ico")) {
			result = "image/x-icon";
		} else if (fileTypeName.equals("jfif") || fileTypeName.equals("jpe")
				|| fileTypeName.equals("jpeg") || fileTypeName.equals("jpg")) {
			result = "image/jpeg";
		} else if (fileTypeName.equals("net")) {
			result = "image/pnetvue";
		} else if (fileTypeName.equals("png") || fileTypeName.equals("bmp")) {
			result = "image/png";
		} else if (fileTypeName.equals("rp")) {
			result = "image/vnd.rn-realpix";
		}

		return result;
	}



	/**
	 * 获取图片后缀
	 *
	 * @param contentType
	 * @return
	 */
	public static String getImageExtension(String contentType) {
		String result = "jpg";
		//http://tools.jb51.net/table/http_content_type
		if (contentType.equals("image/tiff")) {
			result = "tiff";
		} else if (contentType.equals("image/fax")) {
			result = "fax";
		} else if (contentType.equals("image/gif")) {
			result = "gif";
		} else if (contentType.equals("image/x-icon")) {
			result = "ico";
		} else if (contentType.equals("image/jpeg")) {
			result = "jpg";
		} else if (contentType.equals("image/pnetvue")) {
			result = "net";
		} else if (contentType.equals("image/png")) {
			result = "png";
		} else if (contentType.equals("image/vnd.rn-realpix")) {
			result = "rp";
		}

		return result;
	}



	/**
	 * base64转inputStream
	 *
	 * @param base64string
	 * @return
	 */
	private static InputStream BaseToInputStream(String base64string) {
		ByteArrayInputStream stream = null;
		try {
			BASE64Decoder decoder = new BASE64Decoder();
			byte[] bytes1 = decoder.decodeBuffer(base64string);
			stream = new ByteArrayInputStream(bytes1);
		} catch (Exception e) {
			logger.debug("图片BaseToInputStream Exception：{}", e.getMessage());
		}
		return stream;
	}



	public static String getFileSuffix(String srcRealPath) {
		int lastIndex = srcRealPath.lastIndexOf(".");
		String suffix = srcRealPath.substring(lastIndex + 1);
		//        String suffix = srcRealPath.substring(srcRealPath.indexOf(".")+1);
		return suffix;
	}



	/**
	 * base64读取图片像素
	 *
	 * @param imageFilePath
	 * @param height
	 * @param width
	 * @return
	 */
	private static String generateImageBodyBlockStyleAttrByBase64(String imageFilePath, float height, float width) {
		StringBuilder sb = new StringBuilder();
		try {
			try (InputStream input = BaseToInputStream(imageFilePath);) {
				sb = getImageInfo(input, height, width);
			} catch (FileNotFoundException e) {
				logger.debug("安全周报图片错误FileNotFoundException： {}", e.getMessage());
			} catch (IOException e) {
				logger.debug("安全周报图片错误IOException： {}", e.getMessage());
			}
		} catch (Exception e) {
			logger.debug("安全周报 generateImageBodyBlockStyleAttr Exception： {}", e.getMessage());
		}

		return sb.toString();
	}



	private static StringBuilder getImageInfo(InputStream input, float height, float width) throws IOException {
		StringBuilder sb = new StringBuilder();
		BufferedImage sourceImg = ImageIO.read(input);
		if (height == 0) {
			height = sourceImg.getHeight();
		}
		if (width == 0) {
			width = sourceImg.getWidth();
		}
		//将像素转化成pt
		BigDecimal heightValue = new BigDecimal(height * 12 / 16);
		heightValue = heightValue.setScale(2, BigDecimal.ROUND_HALF_UP);
		BigDecimal widthValue = new BigDecimal(width * 12 / 16);
		widthValue = widthValue.setScale(2, BigDecimal.ROUND_HALF_UP);

		sb.append("height:" + heightValue + "pt;");
		sb.append("width:" + widthValue + "pt;");
		sb.append("visibility:visible;");
		sb.append("mso-wrap-style:square; ");
		input.close();
		return sb;
	}

}