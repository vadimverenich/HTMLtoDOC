package com.example;

import org.apache.http.client.fluent.Request;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class HtmlToDocxConverter {

    public static void main(String[] args) throws IOException {
        if (args.length != 2) {
            System.out.println("Usage: java HtmlToDocxConverter <URL> <output-filename>");
            return;
        }

        String url = args[0];
        String outputFilename = args[1];

        // Загрузка HTML содержимого с сайта
        String htmlContent = Request.Get(url).execute().returnContent().asString();

        // Парсинг HTML содержимого
        Document htmlDoc = Jsoup.parse(htmlContent);

        // Создание нового документа DOCX
        XWPFDocument docx = new XWPFDocument();

        // Обработка HTML содержимого и добавление в DOCX
        processHtmlElement(docx, htmlDoc.body());

        // Сохранение DOCX файла
        try (FileOutputStream out = new FileOutputStream(new File(outputFilename))) {
            docx.write(out);
        }

        System.out.println("Document saved as " + outputFilename);
    }

    private static void processHtmlElement(XWPFDocument docx, Element element) {
        for (Element child : element.children()) {
            if (child.tagName().equals("p")) {
                XWPFParagraph paragraph = docx.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(child.text());
            } else if (child.tagName().equals("h1")) {
                XWPFParagraph paragraph = docx.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setFontSize(20);
                run.setText(child.text());
            } else if (child.tagName().equals("h2")) {
                XWPFParagraph paragraph = docx.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setFontSize(18);
                run.setText(child.text());
            } else if (child.tagName().equals("ul")) {
                XWPFParagraph paragraph = docx.createParagraph();
                XWPFRun run = paragraph.createRun();
                Elements items = child.getElementsByTag("li");
                for (Element item : items) {
                    run.setText("- " + item.text());
                    run.addBreak();
                }
            } else if (child.tagName().equals("div") && "ltr".equals(child.attr("dir"))) {
                XWPFParagraph paragraph = docx.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(child.text());
            }
            // Рекурсивно обрабатывать дочерние элементы
            processHtmlElement(docx, child);
        }
    }
}
