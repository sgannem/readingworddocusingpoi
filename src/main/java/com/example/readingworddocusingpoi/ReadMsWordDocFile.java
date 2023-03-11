package com.example.readingworddocusingpoi;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class ReadMsWordDocFile {

    public static void main(String[] args) throws IOException {
        String fileName = "src/main/resources/1_to_106_Mind.docx";
        fetchDataFromFile(fileName);
    }

    private static void fetchDataFromFile(String fileName) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get(fileName)))) {
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(doc);
            String docText = xwpfWordExtractor.getText();
//            System.out.println(docText);
//            String docMetaData = xwpfWordExtractor.getExtendedProperties().
            List<IBodyElement> list = doc.getBodyElements();
            for (IBodyElement temp : list) {
                if (temp instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) temp;
//                    System.out.println("#paragraph:"+paragraph);
//                    System.out.println("firstLineIndent:"+paragraph.getFirstLineIndent());
//                    System.out.println("Text:"+paragraph.getText());
//                    System.out.println("alignment value:"+paragraph.getAlignment().getValue());
//                    System.out.println("runs:"+paragraph.getRuns());

                    System.out.println("text:" + paragraph.getText());
                    System.out.println("alignment:" + paragraph.getAlignment());
                    System.out.println("runs size:" + paragraph.getRuns().size());
                    System.out.println("================================");
                    int cnt = 0;
                    for (XWPFRun run1 : paragraph.getRuns()) {
//                        System.out.println("run.text:" + run1.getText(0));
                        System.out.println("^^^^^^^^^^^^^^^^^^^^^^^^^^^^");
                        System.out.println("run doc - "+(cnt++));
                        System.out.println("run.style:" + run1.getStyle());
                        System.out.println("run.fontName:" + run1.getFontName());
                        System.out.println("run.isBold:" + run1.isBold());
                        System.out.println("run.isItalic:" + run1.isItalic());
                    }
                    System.out.println("================================");
                    System.out.println("style:" + paragraph.getStyle());
                    // Returns numbering format for this paragraph, eg bullet or lowerLetter.
                    System.out.println("numFmt:" + paragraph.getNumFmt());
                    System.out.println("alignment:" + paragraph.getAlignment());
                    System.out.println("isWorldWrapped:" + paragraph.isWordWrapped());
                    System.out.println("--------------------");
                } else if (temp instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) temp;
                    for (XWPFTableRow row : table.getRows()) {
                        System.out.println("row:" + row);
                    }
                }
            }
        }
    }
}
