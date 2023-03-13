package com.example.readingworddocusingpoi;

import lombok.extern.slf4j.Slf4j;
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

@Slf4j
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
//            log.info(docText);
//            String docMetaData = xwpfWordExtractor.getExtendedProperties().
            List<IBodyElement> list = doc.getBodyElements();
            for (IBodyElement temp : list) {
                if (temp instanceof XWPFParagraph) {
                    XWPFParagraph paragraph = (XWPFParagraph) temp;
//                    log.info("#paragraph:"+paragraph);
//                    log.info("firstLineIndent:"+paragraph.getFirstLineIndent());
//                    log.info("Text:"+paragraph.getText());
//                    log.info("alignment value:"+paragraph.getAlignment().getValue());
//                    log.info("runs:"+paragraph.getRuns());

                    log.info("text:" + paragraph.getText());
                    log.info("alignment:" + paragraph.getAlignment());
                    log.info("runs size:" + paragraph.getRuns().size());
                    log.info("================================");
                    int cnt = 0;
                    for (XWPFRun run1 : paragraph.getRuns()) {
//                        log.info("run.text:" + run1.getText(0));
                        log.info("^^^^^^^^^^^^^^^^^^^^^^^^^^^^");
                        log.info("run doc - "+(cnt++));
                        log.info("run.style:" + run1.getStyle());
                        log.info("run.fontName:" + run1.getFontName());
                        log.info("run.isBold:" + run1.isBold());
                        log.info("run.isItalic:" + run1.isItalic());
                    }
                    log.info("================================");
                    log.info("style:" + paragraph.getStyle());
                    // Returns numbering format for this paragraph, eg bullet or lowerLetter.
                    log.info("numFmt:" + paragraph.getNumFmt());
                    log.info("alignment:" + paragraph.getAlignment());
                    log.info("isWorldWrapped:" + paragraph.isWordWrapped());
                    log.info("--------------------");
                } else if (temp instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) temp;
                    for (XWPFTableRow row : table.getRows()) {
                        log.info("row:" + row);
                    }
                }
            }
        }
    }
}
