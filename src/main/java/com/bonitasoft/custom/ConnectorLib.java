package com.bonitasoft.custom;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by Fabrice.R on 20/05/2015.
 *
 * Rules : use bookmark each bookmark must contain the same text as the name of the bookmark (ex : bookmark name : "town", bookmark contain text : "town"
 *
 */
public class ConnectorLib {

    private static final Logger uilLogger = Logger.getLogger("com.bonitasoft.groovy");

    public static void trace(String message){
        try {
            uilLogger.info(message);
            System.out.println(message);
        }catch (Exception ex) {
            uilLogger.severe("trace - Error : " + ex);
        }
    }

    public static String getDateTimeStr(){
        try {
            Calendar calendar = Calendar.getInstance();
            java.util.Date currentDate = calendar.getTime();
            java.sql.Date dateReturn = new java.sql.Date(currentDate.getTime());
            SimpleDateFormat formater = new SimpleDateFormat("yyyyMMddHHmmss");
            return formater.format(dateReturn);
        }catch (Exception ex) {
            trace("getDateTimeStr - Error : " + ex);
            return null;
        }
    }

    public static String sayHello(String who){
        String retour = "";
        try {
            retour = "Hello " + who;

            return retour;
        }catch (Exception ex) {
            trace("sayHello - Exception : " + ex);
            return null;
        }
    }

    public static boolean runConnector(String from, String to, String tmp, String fileName, String fileNameFinale, List listMapping){
        try {
            boolean retour = true;
            Pattern pattern;
            Matcher matcher;
            File file;

            //test fileName
            pattern = Pattern.compile("((\\.(?i)(docx))$)");
            matcher = pattern.matcher(fileName);
            retour = matcher.find();
            if (!retour || (fileName.length() < 2)) {
                trace("error fileName not correct docx expected : " + fileName);
                return false;
            }

            //test fileNameFinale
            pattern = Pattern.compile("((\\.(?i)(pdf))$)");
            matcher = pattern.matcher(fileNameFinale);
            retour = matcher.find();
            if (!retour || (fileName.length() < 2)) {
                trace("error fileNameFinale not correct pdf expected : " + fileName);
                return false;
            }

            //test fileNameFinale
            file = new File(from + fileName);
            if (!file.exists()) {
                trace("error file input not exist : " + from + fileName);
                return false;
            }

            String dateTimeStr = getDateTimeStr();
            String tmpFile = dateTimeStr+"_tmp_"+fileName;

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(from + fileName));

            if(!changeValueBookmark(wordMLPackage, listMapping)){
                trace("error while change bookmarks");
                return false;
            }

            wordMLPackage.save(new File(tmp + tmpFile));

            trace("File tmp genereted : " + tmp + tmpFile);

            retour = docxToPdf(tmp, to, tmpFile, fileNameFinale);

            if(!retour){
                return false;
            }

            trace("File genereted : " + to + fileNameFinale);

            return retour;
        } catch (Exception ex) {
            trace("runConnector - Exception : " + ex);
            return false;
        }
    }

    public static boolean changeValueBookmark(WordprocessingMLPackage wordMLPackage, List listMapping){
        try {
            for (Object elt : listMapping) {
                String key = (String) ((List<String>) elt).get(0);
                String value = (String) ((List<String>) elt).get(1);
                List<Object> parentNode = findContent(wordMLPackage, key);

                //the parent node can contain many different things, sometimes even empty texts.
                for (Object eltInParent : parentNode) {
                    boolean gardian = false;
                    if (eltInParent instanceof org.docx4j.wml.R) {
                        org.docx4j.wml.R theR = (org.docx4j.wml.R) eltInParent;

                        for (Object ob : theR.getContent()) {
                            Object content = ((JAXBElement) ob).getValue();
                            if (content instanceof org.docx4j.wml.Text) {
                                String valueold = ((Text) content).getValue();
                                //We need to choose the good text to edit, we use the key to recognize it
                                if (valueold.equalsIgnoreCase(key)) {
                                    ((Text) content).setValue(value);
                                    gardian = true;
                                    break;
                                }
                            }
                        }

                        if(gardian){
                            break;
                        }
                    }
                }
            }

            return true;
        }catch (Exception ex) {
            trace("changeValueBookmark - changeValueBookmark : " + ex);
            return false;
        }
    }

    public static boolean copyFile(String From, String To, String fileName){
        boolean boolReturn = false;
        try {
            Path FROM = Paths.get(From+fileName);
            Path TO = Paths.get(To + fileName);

            //overwrite existing file, if exists
            CopyOption[] options = new CopyOption[]{
                StandardCopyOption.REPLACE_EXISTING,
                StandardCopyOption.COPY_ATTRIBUTES
            };
            Path returnPath = Files.copy(FROM, TO, options);

            boolReturn = TO.equals(returnPath);

            return boolReturn;
        }catch (Exception ex) {
            trace("copyFile - Exception : " + ex);
            return false;
        }
    }

    public static boolean docxToPdf(String From, String To, String fileNameDocx, String fileNamePdf){
        boolean boolReturn = false;
        try {
            //https://code.google.com/p/xdocreport/wiki/XWPFConverterPDFViaIText
            // 1) Load DOCX into XWPFDocument
            InputStream in = new FileInputStream(new File(From + fileNameDocx));
            XWPFDocument document = new XWPFDocument(in);

            // 2) Prepare Pdf options
            // iText during conversion process has to access fonts with proper encoding. Otherwise some
            // diactric characters may be not converted properly (missing from the resulting pdf)
            // The converter uses underlying operating system encoding as the default value.
            // It is also possible to set it explicitely
            //PdfOptions options = PdfOptions.create().fontEncoding("windows-1250");
            PdfOptions options = PdfOptions.create();

            // 3) Convert XWPFDocument to Pdf
            OutputStream out = new FileOutputStream(To + fileNamePdf);
            PdfConverter.getInstance().convert(document, out, options);

            File f = new File(To + fileNamePdf);

            boolReturn = ((f.exists()) && (f.length() > 0));

            return boolReturn;
        }catch (Exception ex) {
            trace("docxToPdf - Exception : " + ex);
            return false;
        }
    }

    public static WordprocessingMLPackage changeBookmarkWithImg(WordprocessingMLPackage wordMLPackage, String img){
        try {
            File file = new File(img);
            byte[] bytes = convertImageToByteArray(file);
            addImageToPackage(wordMLPackage, bytes);

            return wordMLPackage;
        }catch (Exception ex) {
            trace("changeBookmarkWithImg - Exception : " + ex);
            return null;
        }
    }

    public static List<Object> findContent(WordprocessingMLPackage wordMLPackage, String key){
        try {
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            List<Object> paragraphs = documentPart.getContent();

            RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
            new TraversalUtil(paragraphs, rt);

            for (CTBookmark bm : rt.getStarts()) {
                String name = bm.getName();
                if (name.equalsIgnoreCase(key)) {
                    return ((ContentAccessor)(bm.getParent())).getContent();
                }
            }

            return null;
        }catch (Exception ex) {
            trace("findContent - Exception : " + ex);
            return null;
        }
    }

    /**
     *  Docx4j contains a utility method to create an image part from an array of
     *  bytes and then adds it to the given package. In order to be able to add this
     *  image to a paragraph, we have to convert it into an inline object. For this
     *  there is also a method, which takes a filename hint, an alt-text, two ids
     *  and an indication on whether it should be embedded or linked to.
     *  One id is for the drawing object non-visual properties of the document, and
     *  the second id is for the non visual drawing properties of the picture itself.
     *  Finally we add this inline object to the paragraph and the paragraph to the
     *  main document of the package.
     *
     *  @param wordMLPackage The package we want to add the image to
     *  @param bytes         The bytes of the image
     *  @throws Exception    Sadly the createImageInline method throws an Exception
     *                       (and not a more specific exception type)
     */
    private static void addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        int docPrId = 1;
        int cNvPrId = 2;
        Inline inline = imagePart.createImageInline("Filename hint","Alternative text", docPrId, cNvPrId, false);

        P paragraph = addInlineImageToParagraph(inline);

        wordMLPackage.getMainDocumentPart().addObject(paragraph);
    }

    /**
     *  We create an object factory and use it to create a paragraph and a run.
     *  Then we add the run to the paragraph. Next we create a drawing and
     *  add it to the run. Finally we add the inline object to the drawing and
     *  return the paragraph.
     *
     * @param   inline The inline object containing the image.
     * @return  the paragraph containing the image
     */
    private static P addInlineImageToParagraph(Inline inline) {
        // Now add the in-line image to a paragraph
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        paragraph.getContent().add(run);
        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return paragraph;
    }

    /**
     * Convert the image from the file into an array of bytes.
     *
     * @param file  the image file to be converted
     * @return      the byte array containing the bytes from the image
     * @throws FileNotFoundException
     * @throws IOException
     */
    private static byte[] convertImageToByteArray(File file) throws FileNotFoundException, IOException {
        InputStream is = new FileInputStream(file );
        long length = file.length();
        // You cannot create an array using a long, it needs to be an int.
        if (length > Integer.MAX_VALUE) {
            System.out.println("File too large!!");
        }
        byte[] bytes = new byte[(int)length];
        int offset = 0;
        int numRead = 0;
        while (offset < bytes.length && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
            offset += numRead;
        }
        // Ensure all the bytes have been read
        if (offset < bytes.length) {
            System.out.println("Could not completely read file " +file.getName());
        }
        is.close();
        return bytes;
    }
}