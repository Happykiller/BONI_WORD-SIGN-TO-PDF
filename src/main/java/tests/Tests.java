package tests;

import com.bonitasoft.ConnectorLib;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Fabrice.R on 19/01/2015.
 */
public class Tests {
    private static boolean DELETE_BOOKMARK = true;

    private static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();

    public static void main(String [ ] args){
        ConnectorLib.trace("===== Mes tests ======");
        testSayHello();
        //testCopyFile();
        testEditWord();
    }

    public static void testSayHello(){
        String attendu = "Hello you";

        String retour = ConnectorLib.sayHello("you");

        ConnectorLib.trace("Test testSayHello : " + attendu.equals(retour));
    }

    public static void testEditWord(){
        Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
        map.put( new DataFieldName("CodePostal"), "whale shark");


        WordprocessingMLPackage wordMLPackage = null;
        try {
            wordMLPackage = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "\\resources\\build\\patern-signets.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // Before..
        // System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));

        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart.getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        try {
            replaceBookmarkContents(body.getContent(), map);
        } catch (Exception e) {
            e.printStackTrace();
        }

        // After
        // System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));

        // save the docx...
        try {
            wordMLPackage.save(new File(System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        }

        /*org.docx4j.convert.out.pdf.PdfConversion c
//				= new org.docx4j.convert.out.pdf.viaHTML.Conversion(wordMLPackage);
                = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(wordMLPackage);
//				= new org.docx4j.convert.out.pdf.viaIText.Conversion(wordMLPackage);

        ((org.docx4j.convert.out.pdf.viaXSLFO.Conversion)c).setSaveFO(new java.io.File(System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.fo"));

        OutputStream os = null;
        try {
            os = new java.io.FileOutputStream(System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.pdf");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            c.output(os, new PdfSettings() );
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        System.out.println("Saved " + System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.pdf");*/

        //https://code.google.com/p/xdocreport/wiki/XWPFConverterPDFViaIText
        // 1) Load DOCX into XWPFDocument
        InputStream in= null;
        try {
            in = new FileInputStream(new File(System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.docx"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XWPFDocument document = null;
        try {
            document = new XWPFDocument(in);
        } catch (IOException e) {
            e.printStackTrace();
        }

// 2) Prepare Pdf options
        // iText during conversion process has to access fonts with proper encoding. Otherwise some
// diactric characters may be not converted properly (missing from the resulting pdf)
// The converter uses underlying operating system encoding as the default value.
// It is also possible to set it explicitely
        //PdfOptions options = PdfOptions.create().fontEncoding("windows-1250");
        PdfOptions options = PdfOptions.create();

// 3) Convert XWPFDocument to Pdf
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File(System.getProperty("user.dir") + "\\resources\\build\\OUT_BookmarksTextInserter.pdf"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void replaceBookmarkContents(List<Object> paragraphs, Map<DataFieldName, String> data) throws Exception {

        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {

            // do we have data for this one?
            if (bm.getName()==null) continue;
            String value = data.get(new DataFieldName(bm.getName()));
            if (value==null) continue;

            try {
                // Can't just remove the object from the parent,
                // since in the parent, it may be wrapped in a JAXBElement
                List<Object> theList = null;
                if (bm.getParent() instanceof P) {
                    System.out.println("OK!");
                    theList = ((ContentAccessor)(bm.getParent())).getContent();
                } else {
                    continue;
                }

                int rangeStart = -1;
                int rangeEnd=-1;
                int i = 0;
                for (Object ox : theList) {
                    Object listEntry = XmlUtils.unwrap(ox);
                    if (listEntry.equals(bm)) {
                        if (DELETE_BOOKMARK) {
                            rangeStart=i;
                        } else {
                            rangeStart=i+1;
                        }
                    } else if (listEntry instanceof  CTMarkupRange) {
                        if ( ((CTMarkupRange)listEntry).getId().equals(bm.getId())) {
                            if (DELETE_BOOKMARK) {
                                rangeEnd=i;
                            } else {
                                rangeEnd=i-1;
                            }
                            break;
                        }
                    }
                    i++;
                }

                if (rangeStart>0 && rangeEnd>rangeStart) {

                    // Delete the bookmark range
                    for (int j =rangeEnd; j>=rangeStart; j--) {
                        theList.remove(j);
                    }

                    // now add a run
                    org.docx4j.wml.R  run = factory.createR();
                    org.docx4j.wml.Text  t = factory.createText();
                    run.getContent().add(t);
                    t.setValue(value);

                    theList.add(rangeStart, run);
                }

            } catch (ClassCastException cce) {
                System.out.println(cce.getMessage());
            }
        }


    }
}