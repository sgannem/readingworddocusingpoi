# readingworddocusingpoi
Reading ms word document file using apache poi library

## Read Each Paragraph of a Word Document
Among the many methods defined in XWPFDocument class, we can use getParagraphs() to read a .docx word document paragraph wise.This method returns a list of all the paragraphs(XWPFParagraph) of a word document. Again the XWPFParagraph has many utils method defined to extract information related to any paragraph such as text alignment, style associated with the paragrpahs.

### sample snippet
````
FileInputStream fis = new FileInputStream("test.docx");
XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
List paragraphList = xdoc.getParagraphs();
// reading each paragraph
for (XWPFParagraph paragraph : paragraphList) {
	log.info(paragraph.getText());
	log.info(paragraph.getAlignment());
	System.out.print(paragraph.getRuns().size());
	log.info(paragraph.getStyle());

	// Returns numbering format for this paragraph, eg bullet or lowerLetter.
	log.info(paragraph.getNumFmt());
	log.info(paragraph.getAlignment());

	log.info(paragraph.isWordWrapped());

	log.info("********************************************************************");
}
````

### how to run the project
#### pre-requisites 
* jdk 1.8
* maven 3.0.5 and above
#### how to run project
```
git clone <project-git-hub-url>
cd <cloned-project>
mvn clean install
java -Dlog4j2.contextSelector=org.apache.logging.log4j.core.async.AsyncLoggerContextSelector -jar target/readingworddocusingpoi-0.0.1-SNAPSHOT.jar
```
#### sample output
```
--------------------
text:MIND
alignment:CENTER
runs size:1
================================
^^^^^^^^^^^^^^^^^^^^^^^^^^^^
run doc - 0
run.style:
run.fontName:Times New Roman
run.isBold:true
run.isItalic:false
================================
style:null
numFmt:null
alignment:CENTER
isWorldWrapped:false
--------------------
text:ABANDONED. (See Forsaken.)
alignment:BOTH
runs size:2
================================
```

## References
* https://www.devglan.com/corejava/parsing-word-document-example