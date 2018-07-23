package com.fidel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;

public class Main {
	private static final Logger log = LoggerFactory.getLogger(Main.class);

	public static void main(String[] args) throws IOException {
		Document doc = Jsoup.connect("https://en.wikipedia.org/wiki/Main_Page/").get();
		writeExcelFile(doc);

/*		try (Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("Output.txt"), "utf-8"))) {
			Document doc = Jsoup.connect("https://en.wikipedia.org/wiki/Main_Page/").get();
			log.info(doc.title());
			writer.write(doc.title() + System.lineSeparator());
			Elements newsHeadlines = doc.select("#mp-itn b a");
			for (Element headline : newsHeadlines) {
				log.info("{}\n\t{}", headline.attr("title"), headline.absUrl("href"));
				writer.write(headline.attr("title") + ";" + headline.absUrl("href") + System.lineSeparator());
			}
		} catch (IOException ex) {
			log.error("IOException");
		}*/
	}

	private static void writeExcelFile(Document doc) throws IOException {
		Workbook wb = new HSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("New sheet");
		Elements newsHeadlines = doc.select("#mp-itn b a");
		int i = 0;
		for (Element headline : newsHeadlines) {
			Row row = sheet.createRow((short)i);
			row.createCell(0).setCellValue(createHelper.createRichTextString(headline.attr("title")));
			row.createCell(1).setCellValue(createHelper.createRichTextString(headline.absUrl("href")));
			i++;
		}

		FileOutputStream fileOut = new FileOutputStream(doc.title()+".xls");
		wb.write(fileOut);
		fileOut.close();
	}

}


