package com.alex.common.utils.freemarker;

import freemarker.core.ParseException;
import freemarker.template.Configuration;
import freemarker.template.MalformedTemplateNameException;
import freemarker.template.Template;
import freemarker.template.TemplateNotFoundException;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * @version V1.0
 * @Description:word导出帮助类 通过freemarker模板引擎来实现
 * @author:LiaoFei
 * @date :2016-3-24 下午3:49:25
 */
public class WordGeneratorWithFreemarker {
	private static Configuration         configuration = null;
	private static Map<String, Template> allTemplates  = null;

	static {
		configuration = new Configuration(Configuration.VERSION_2_3_23);
		configuration.setDefaultEncoding("utf-8");
		configuration.setClassicCompatible(true);
		configuration.setClassForTemplateLoading(WordGeneratorWithFreemarker.class, "/static/ftl");
	}

	private WordGeneratorWithFreemarker() {

	}



	public static void createDoc(Map<String, Object> dataMap, String templateName, OutputStream out)
			throws TemplateNotFoundException, MalformedTemplateNameException, ParseException, IOException {

		Template t = configuration.getTemplate(templateName);

		WordHtmlGeneratorHelper.handleAllObject(dataMap);

		try {
			// 这个地方不能使用FileWriter因为需要指定编码类型否则生成的Word文档会因为有无法识别的编码而无法打开
			Writer w = new OutputStreamWriter(out);
			t.process(dataMap, w);
			w.close();
		} catch (Exception ex) {
			ex.printStackTrace();
			throw new RuntimeException(ex);
		}
		return;
	}



	public static void main(String[] args) {
		HashMap<String, Object> data = new HashMap<String, Object>();

		data.put("dwmc", WordHtmlGeneratorHelper.string2Ascii("湖南省交通科学研究 院"));
		data.put("picture", WordHtmlGeneratorHelper.string2Ascii("这里显示图片"));

		String docFilePath = "d:\\temp" + (int) (Math.random() * 100000) + ".doc";
		File f = new File(docFilePath);
		OutputStream out;
		try {
			out = new FileOutputStream(f);
			WordGeneratorWithFreemarker.createDoc(data, "xmht", out);

		} catch (FileNotFoundException e) {

		} catch (MalformedTemplateNameException e) {
			e.printStackTrace();
		} catch (ParseException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}  