import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;


import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;

public class Main {

  static Dimension getImageDimension(File imgFile) throws IOException {
    BufferedImage img = ImageIO.read(imgFile);
    return new Dimension(img.getWidth(), img.getHeight());
  }

  public static void main(String[] args) throws Exception {

    //Blank Document
    XWPFDocument document = new XWPFDocument();
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
    XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.FIRST);

    headerParagraph = header.createParagraph();
    headerParagraph.setAlignment(ParagraphAlignment.BOTH);

    run = headerParagraph.createRun();
    run.setText("The Header:");

    File imgFile = new File("C:\\Users\\ucha.chaduneli\\IdeaProjects\\poiDoc\\src\\main\\resources\\background.png");
//    InputStream headerImg = new FileInputStream("C:\\Users\\ucha.chaduneli\\IdeaProjects\\poiDoc\\src\\main\\resources\\background.png");
//    paragraph.createRun().addPicture(headerImg, Document.PICTURE_TYPE_PNG, "background.png", Units.toEMU(500), Units.toEMU(800));
    Dimension dim = getImageDimension(imgFile);
    double width = dim.getWidth();
    double height = dim.getHeight();

    double scaling = 1.0;
    if (width > 80 * 7) scaling = (80 * 7) / width; //scale width not to be greater than 6 inches
    InputStream in = new FileInputStream(imgFile);
    paragraph.createRun().addPicture(in, Document.PICTURE_TYPE_PNG, "background.png",
            Units.toEMU(width * scaling), Units.toEMU(height * scaling));
    in.close();

    run = paragraph.createRun();
    run.addBreak(BreakType.PAGE);
    run.addBreak(BreakType.TEXT_WRAPPING);
    run.setText("Main Text babli bubli adlkasj;la doai;sjd as;dj asl;djas ldjald jaslj\n");

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
      Desktop.getDesktop().open(new File("C:\\Users\\ucha.chaduneli\\Desktop\\myDoc.docx"));
    }
  }
}
