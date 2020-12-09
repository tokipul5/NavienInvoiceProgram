import java.io.*;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation.Enum;


public class CopyWord {
    private String pathTemplate;
    private String pathNewDoc;
    public CopyWord(String pathTemplate, String pathNewDoc) {
        this.pathTemplate = pathTemplate;
        this.pathNewDoc = pathNewDoc;
    }
    public void copy() throws IOException {
        XWPFDocument srcDoc = new XWPFDocument(new FileInputStream(pathTemplate));

        CustomXWPFDocument destDoc = new CustomXWPFDocument();

        // Copy document layout.
        copyLayout(srcDoc, destDoc);

        OutputStream out = new FileOutputStream(pathNewDoc);

        for (IBodyElement bodyElement : srcDoc.getBodyElements()) {

            BodyElementType elementType = bodyElement.getElementType();

            if (elementType == BodyElementType.PARAGRAPH) {

                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;

                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(srcPr.getStyleID()));

                boolean hasImage = false;

                XWPFParagraph dstPr = destDoc.createParagraph();

                // Extract image from source docx file and insert into destination docx file.
                for (XWPFRun srcRun : srcPr.getRuns()) {

                    // You need next code when you want to call XWPFParagraph.removeRun().
                    dstPr.createRun();

                    if (srcRun.getEmbeddedPictures().size() > 0)
                        hasImage = true;

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                        byte[] img = pic.getPictureData().getData();

                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();

                        try {
                            // Working addPicture Code below...
                            String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),
                                    Document.PICTURE_TYPE_PNG);
                            destDoc.createPictureCxCy(blipId, destDoc.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
                                    cx, cy);

                        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
                            e.printStackTrace();
                        }
                    }
                }

                if (hasImage == false)
                {
                    int pos = destDoc.getParagraphs().size() - 1;
                    destDoc.setParagraph(srcPr, pos);
                }

            } else if (elementType == BodyElementType.TABLE) {

                XWPFTable table = (XWPFTable) bodyElement;

                copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(table.getStyleID()));

                destDoc.createTable();

                int pos = destDoc.getTables().size() - 1;

                destDoc.setTable(pos, table);
            }
        }

        destDoc.write(out);
        out.close();
        destDoc.close();
    }
    // Copy Styles of Table and Paragraph.
    private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style)
    {
        if (destDoc == null || style == null)
            return;

        if (destDoc.getStyles() == null) {
            destDoc.createStyles();
        }

        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
        for (XWPFStyle xwpfStyle : usedStyleList) {
            destDoc.getStyles().addStyle(xwpfStyle);
        }
    }

    // Copy Page Layout.
    //
    // if next error message shows up, download "ooxml-schemas-1.1.jar" file and
    // add it to classpath.
    //
    // [Error]
    // The type org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar
    // cannot be resolved.
    // It is indirectly referenced from required .class files
    //
    // This error happens because there is no CTPageMar class in
    // poi-ooxml-schemas-3.10.1-20140818.jar.
    //
    // [ref.] http://poi.apache.org/faq.html#faq-N10025
    // [ref.] http://poi.apache.org/overview.html#components
    //
    // < ooxml-schemas 1.1 download >
    // http://repo.maven.apache.org/maven2/org/apache/poi/ooxml-schemas/1.1/
    //
    private static void copyLayout(XWPFDocument srcDoc, XWPFDocument destDoc)
    {
        CTPageMar pgMar = srcDoc.getDocument().getBody().getSectPr().getPgMar();

        BigInteger bottom = pgMar.getBottom();
        BigInteger footer = pgMar.getFooter();
        BigInteger gutter = pgMar.getGutter();
        BigInteger header = pgMar.getHeader();
        BigInteger left = pgMar.getLeft();
        BigInteger right = pgMar.getRight();
        BigInteger top = pgMar.getTop();

        CTPageMar addNewPgMar = destDoc.getDocument().getBody().addNewSectPr().addNewPgMar();

        addNewPgMar.setBottom(bottom);
        addNewPgMar.setFooter(footer);
        addNewPgMar.setGutter(gutter);
        addNewPgMar.setHeader(header);
        addNewPgMar.setLeft(left);
        addNewPgMar.setRight(right);
        addNewPgMar.setTop(top);

        CTPageSz pgSzSrc = srcDoc.getDocument().getBody().getSectPr().getPgSz();

        BigInteger code = pgSzSrc.getCode();
        BigInteger h = pgSzSrc.getH();
        Enum orient = pgSzSrc.getOrient();
        BigInteger w = pgSzSrc.getW();

        CTPageSz addNewPgSz = destDoc.getDocument().getBody().addNewSectPr().addNewPgSz();

        addNewPgSz.setCode(code);
        addNewPgSz.setH(h);
        addNewPgSz.setOrient(orient);
        addNewPgSz.setW(w);
    }
}
