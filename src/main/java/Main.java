import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigInteger;

public class Main {

  static Dimension getImageDimension(File imgFile) throws IOException {
    BufferedImage img = ImageIO.read(imgFile);
    return new Dimension(img.getWidth(), img.getHeight());
  }

  public static void main(String[] args) throws Exception {

    //Blank Document
    XWPFDocument document = new XWPFDocument();

    CTBody body = document.getDocument().getBody();
    if (!body.isSetSectPr()) {
      body.addNewSectPr();
    }

    CTSectPr section = body.getSectPr();
    if (!section.isSetPgSz()) {
      section.addNewPgSz();
    }

    CTPageSz pageSize = section.getPgSz();
    pageSize.setOrient(STPageOrientation.LANDSCAPE);
//A4 = 595x842 / multiply 20 since BigInteger represents 1/20 Point
    pageSize.setW(BigInteger.valueOf(16840));
    pageSize.setH(BigInteger.valueOf(11900));

    CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
    CTPageMar pageMar = sectPr.addNewPgMar();
    pageMar.setLeft(BigInteger.valueOf(0L));
    pageMar.setTop(BigInteger.valueOf(0L));
    pageMar.setRight(BigInteger.valueOf(0L));
    pageMar.setBottom(BigInteger.valueOf(0L));

    //Write the Document in file system
    FileOutputStream out = new FileOutputStream(new File("C:\\Users\\home\\Desktop\\myDoc.docx"));

    //create Paragraph
    XWPFParagraph paragraph = document.createParagraph();
    XWPFParagraph headerParagraph = document.createParagraph();
    XWPFRun run;

    // create header-footer
    XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
    if (headerFooterPolicy == null) headerFooterPolicy = document.createHeaderFooterPolicy();

    // create header start
    XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.FIRST);

    headerParagraph = header.createParagraph();
    headerParagraph.setAlignment(ParagraphAlignment.CENTER);

    run = headerParagraph.createRun();
    run.setText("The Header:");

    File imgFile = new File("C:\\Users\\home\\IdeaProjects\\poi\\src\\main\\resources\\background.png");
    Dimension dim = getImageDimension(imgFile);
    double width = dim.getWidth();
    double height = dim.getHeight();

    double scaling = 1.0;
    if (width > 82 * 10.3) scaling = (82 * 10.3) / width; //scale width not to be greater than 6 inches
    InputStream in = new FileInputStream(imgFile);
    paragraph.setAlignment(ParagraphAlignment.BOTH);
    paragraph.createRun().addPicture(in, Document.PICTURE_TYPE_PNG, "background.png",
            Units.toEMU(width * scaling), Units.toEMU(height * scaling));
    in.close();


//    textbox
//    CTGroup ctGroup = CTGroup.Factory.newInstance();
//    CTShape ctShape = ctGroup.addNewShape();
//    ctShape.setStyle("width:300pt;height:100pt");
//    CTTxbxContent ctTxbxContent = ctShape.addNewTextbox().addNewTxbxContent();
//    ctTxbxContent.addNewP().addNewR().addNewT().setStringValue("The TextBox text...");
//    Node ctGroupNode = ctGroup.getDomNode();
//    CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
//    run=paragraph.createRun();
//    CTR cTR = run.getCTR();
//    cTR.addNewPict();
//    cTR.setPictArray(0, ctPicture);
//    *************


//    run = paragraph.createRun();
//    run.addBreak(BreakType.PAGE);
//    run.addBreak(BreakType.TEXT_WRAPPING);
//    run.setText("Main Text babli bubli adlkasj;la doai;sjd as;dj asl;djas ldjald jaslj\n");

    // create footer start
    XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);

    paragraph = footer.createParagraph();
    paragraph.setAlignment(ParagraphAlignment.LEFT);

    run = paragraph.createRun();
    run.setText("The Footer: ");

    document.write(out);

    //Close document
    out.close();
    System.out.println("Doc Generated successfully");
    if (Desktop.isDesktopSupported()) {
      Desktop.getDesktop().open(new File("C:\\Users\\home\\Desktop\\myDoc.docx"));
    }
  }
}
