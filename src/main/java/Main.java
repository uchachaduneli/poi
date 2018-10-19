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

    // მარჯინების დასმა
//    CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
//    CTPageMar pageMar = sectPr.addNewPgMar();
//    pageMar.setLeft(BigInteger.valueOf(10L));
//    pageMar.setTop(BigInteger.valueOf(10L));
//    pageMar.setRight(BigInteger.valueOf(10L));
//    pageMar.setBottom(BigInteger.valueOf(10L));

    //Write the Document in file system
    FileOutputStream out = new FileOutputStream(new File("C:\\Users\\ucha.chaduneli\\Desktop\\myDoc.docx"));

    //create Paragraph
    XWPFParagraph paragraph = document.createParagraph();
    XWPFParagraph headerParagraph = document.createParagraph();
    XWPFRun run;

    // create header-footer
    XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
    if (headerFooterPolicy == null) headerFooterPolicy = document.createHeaderFooterPolicy();

    // create header start
    XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

    headerParagraph = header.createParagraph();
    headerParagraph.setAlignment(ParagraphAlignment.LEFT);
    run = headerParagraph.createRun();
    run.setText("MIMINO TRAVEL GEORGIA");

    File imgFile = new File("C:\\Users\\ucha.chaduneli\\IdeaProjects\\poi\\src\\main\\resources\\background.png");
    Dimension dim = getImageDimension(imgFile);
    double width = dim.getWidth();
    double height = dim.getHeight();

    double scaling = 1.0;
    if (width > 50 * 7) scaling = (50 * 7) / width; //scale width not to be greater than 6 inches
    if (height > 40 * 7) scaling = (40 * 7) / height;
    InputStream in = new FileInputStream(imgFile);
    paragraph.setAlignment(ParagraphAlignment.CENTER);
    paragraph.createRun().addPicture(in, Document.PICTURE_TYPE_PNG, "background.png",
            Units.toEMU(width * scaling), Units.toEMU(height * scaling));
    in.close();

//    File imgFile = new File("C:\\Users\\ucha.chaduneli\\IdeaProjects\\poi\\src\\main\\resources\\background.png");
//    Dimension dim = getImageDimension(imgFile);
//    double width = dim.getWidth();
//    double height = dim.getHeight();
//
//    double scaling = 1.0;
//    if (width > 82 * 10.3) scaling = (82 * 10.3) / width; //scale width not to be greater than 6 inches
//    InputStream in = new FileInputStream(imgFile);
//    paragraph.setAlignment(ParagraphAlignment.BOTH);
//    paragraph.createRun().addPicture(in, Document.PICTURE_TYPE_PNG, "background.png",
//            Units.toEMU(width * scaling), Units.toEMU(height * scaling));
//    in.close();

    // create footer start
    XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);

    paragraph = footer.createParagraph();
    paragraph.setAlignment(ParagraphAlignment.BOTH);
    paragraph.setBorderTop(Borders.THICK);
    run = paragraph.createRun();
    run.setText("The Footer ტექსტ: ");


    paragraph = document.createParagraph();
    run = paragraph.createRun();
    run.setText("Breakfast at the hotel. Early in the morning start very scenic drive up through the High Caucasus Mountains along the Georgian Military Highway towards Kazbegi region. On the way we will drive across the Cross Pass (2 395m.) and along the Tergi River that brings us to Kazbegi – main town in this region. From the centre of Kazbegi drive by 4x4 through beautiful valleys and woodland leads us to Gergeti Holy Trinity church (14th century), stunningly located on a hilltop (2170 m.) ");
    run.setText("Breakfast at the hotel. Early in the morning start very scenic drive up through the High Caucasus Mountains along the Georgian Military Highway towards Kazbegi region. On the way we will drive across the Cross Pass (2 395m.) and along the Tergi River that brings us to Kazbegi – main town in this region. From the centre of Kazbegi drive by 4x4 through beautiful valleys and woodland leads us to Gergeti Holy Trinity church (14th century), stunningly located on a hilltop (2170 m.) ");
    run.setText("Breakfast at the hotel. Early in the morning start very scenic drive up through the High Caucasus Mountains along the Georgian Military Highway towards Kazbegi region. On the way we will drive across the Cross Pass (2 395m.) and along the Tergi River that brings us to Kazbegi – main town in this region. From the centre of Kazbegi drive by 4x4 through beautiful valleys and woodland leads us to Gergeti Holy Trinity church (14th century), stunningly located on a hilltop (2170 m.) ");

    document.write(out);

    //Close document
    out.close();
    System.out.println("Doc Generated successfully");
    if (Desktop.isDesktopSupported()) {
      Desktop.getDesktop().open(new File("C:\\Users\\ucha.chaduneli\\Desktop\\myDoc.docx"));
    }
  }
}
