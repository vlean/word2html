package com.abc.word2Html.util;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import cn.hutool.core.img.ImgUtil;
import com.sun.deploy.util.StringUtils;
import de.tototec.cmdoption.CmdOption;
import de.tototec.cmdoption.CmdlineParser;
import de.tototec.cmdoption.CmdlineParserException;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;
import sun.applet.Main;

public class Word2Html {
	public static class Config {

		@CmdOption(names = {"--help", "-h"}, description = "Show this help.", isHelp = true)
		public boolean help;

		@CmdOption(names = {"--input", "-i"}, args = {"INPUT"}, description = "input file path, support doc,docx.")
		public String input = "";

		@CmdOption(names = {"--output", "-o"}, args = {"OUTPUT"}, description = "output file path")
		public String output = "";

	}

	public static void main(String[] args) throws Throwable {

		Config config = new Config();
		CmdlineParser cp = new CmdlineParser(config);
		cp.setProgramName("word2html");
		cp.setAboutLine("convert word to html v1.0");

		try {
			cp.parse(args);
		} catch (CmdlineParserException e) {
			System.err.println("Error: " + e.getLocalizedMessage() + "\nRun word2html --help for help.");
			System.exit(1);
		}

		if (config.help) {
			cp.usage();
			System.exit(0);
		}
		toHtml(config.input, config.output);

	}
	
	private static void toHtml(String input,String output) throws Throwable{
		File f = new File(input);
		if(!f.exists()){
			System.out.println("文件不存在");
			return;
		}
		if (!(output.endsWith("html") || output.endsWith("htm"))) {
			System.out.println("输出文件必须为html格式");
			return;
		}
		if(input.endsWith("doc")){
			doc(f, output);
		}else if(input.endsWith("docx")){
			docx(f, output);
		} else {
			System.out.println("不支持的文件格式");
		}
	}
	
	private static String getFileName(File f){
		String fileName = f.getName();
		return fileName.substring(0, !fileName.contains(".docx") ?fileName.indexOf(".doc"):fileName.indexOf(".docx"));
	}
	
	private static void docx(File f, String output) throws Throwable {

		// 生成 XWPFDocument
		InputStream in = new FileInputStream(f);
		XWPFDocument document = new XWPFDocument(in);

		File file1 = new File(output);
		// 禁止图片
		// 准备 XHTML 选项 (设置 IURIResolver，把图片放到文件绝对路径下image/word/media文件夹
		File imageFolderFile = new File(file1.getParentFile().getPath()+"image/"+getFileName(f));
		XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
		options.setExtractor(new FileImageExtractor(imageFolderFile));
		options.setIgnoreStylesIfUnused(false);
		options.setFragment(true);

		// 将XWPFDocument 转换为  XHTML
		OutputStream out = new FileOutputStream(file1);
		XHTMLConverter.getInstance().convert(document, out, options);
	}

	public static  void doc(final File f, String output) throws Throwable{
		// 生成 HWPFDocument 
		InputStream input = new FileInputStream(f);
		HWPFDocument wordDocument = new HWPFDocument(input);
		// 把图片放到文件绝对路径下image/word/media文件夹
		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
				DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
		wordToHtmlConverter.setPicturesManager((content, pictureType, suggestedName, widthInches, heightInches) -> {
			BufferedImage bufferedImage = ImgUtil.toImage(content);
			String base64Img = ImgUtil.toBase64(bufferedImage, pictureType.getExtension());
			//  带图片的word，则将图片转为base64编码，保存在一个页面中
			return "data:;base64," + base64Img;
		});
		wordToHtmlConverter.processDocument(wordDocument);
		Document htmlDocument = wordToHtmlConverter.getDocument();
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		DOMSource domSource = new DOMSource(htmlDocument);
		StreamResult streamResult = new StreamResult(outStream);
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer serializer = tf.newTransformer();
		serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
		serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		serializer.setOutputProperty(OutputKeys.METHOD, "html");
		serializer.transform(domSource, streamResult);
		outStream.close();
		String content = new String(outStream.toByteArray());
		FileUtils.writeStringToFile(new File(output), content, "utf-8");
	}
}