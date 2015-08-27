package tests;

import com.bonitasoft.custom.ConnectorLib;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Fabrice.R on 19/01/2015.
 */
public class Tests {

    public static void main(String [ ] args){
        ConnectorLib.trace("===== Mes tests ======");
        //testSayHello();
        //testListToMapf();
        testRunConnector();
        //testCopyFile();
        //testDocxToPdf();
        //testChangeBookmarkWithString();
        //testInsertImg();
        //testFindContent();
    }

    public static void testSayHello(){
        String attendu = "Hello you";

        String retour = ConnectorLib.sayHello("you");

        ConnectorLib.trace("Test testSayHello : " + attendu.equals(retour));
    }

    public static void testRunConnector(){
        List<List> list = new ArrayList<List>();

        List<String> eltCodePostal = new ArrayList<String>();
        eltCodePostal.add("CodePostal");
        eltCodePostal.add("38100");
        list.add(eltCodePostal);

        List<String> eltVille = new ArrayList<String>();
        eltVille.add("VILLE");
        eltVille.add("Grenoble");
        list.add(eltVille);

        List<String> eltDateJour = new ArrayList<String>();
        eltDateJour.add("dateJour");
        eltDateJour.add("2022-02-22");
        list.add(eltDateJour);

        boolean retour = ConnectorLib.runConnector(
            System.getProperty("user.dir") + "\\resources\\"
            , System.getProperty("user.dir") + "\\resources\\build\\"
            , System.getProperty("user.dir") + "\\resources\\tmp\\"
            , "patern-signets.docx"
            , ConnectorLib.getDateTimeStr()+"_builded.pdf"
            , list
        );

        ConnectorLib.trace("Test testRunConnector : " + retour);
    }

    public static void testCopyFile(){
        boolean retour = ConnectorLib.copyFile(System.getProperty("user.dir") + "\\resources\\", System.getProperty("user.dir") + "\\resources\\build\\", "testCopyFile_"+ConnectorLib.getDateTimeStr()+".docx");

        ConnectorLib.trace("Test testCopyFile : " + retour);
    }

    public static void testDocxToPdf(){
        boolean retour = ConnectorLib.docxToPdf(System.getProperty("user.dir") + "\\resources\\", System.getProperty("user.dir") + "\\resources\\build\\", "patern-signets.docx", "testDocxToPdf_"+ConnectorLib.getDateTimeStr()+".pdf");

        ConnectorLib.trace("Test testDocxToPdf : " + retour);
    }

    public static void testListToMapf(){
        List<List> list = new ArrayList<List>();

        List<String> elt = new ArrayList<String>();
        elt.add("key");
        elt.add("value");

        list.add(elt);

        Map<DataFieldName, String> retour = ConnectorLib.listToMap(list);

        ConnectorLib.trace("Test testListToMapf : " + !retour.isEmpty());
    }

    public static void testChangeBookmarkWithString() {
        try {
            Map<DataFieldName, String> map = new HashMap<DataFieldName, String>();
            map.put(new DataFieldName("CodePostal"), "38100");

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "\\resources\\patern-signets.docx"));

            ConnectorLib.changeBookmarkWithString(wordMLPackage, map);

            wordMLPackage.save(new File(System.getProperty("user.dir") + "\\resources\\build\\testChangeBookmarkWithString_"+ConnectorLib.getDateTimeStr()+".docx"));

            ConnectorLib.trace("Test testChangeBookmarkWithString : " + true);
        }catch (Exception ex) {
            ConnectorLib.trace("testChangeBookmarkWithString - Exception : " + ex);
        }
    }

    public static void testInsertImg(){
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "\\resources\\patern-signets.docx"));

            ConnectorLib.changeBookmarkWithImg(wordMLPackage, System.getProperty("user.dir") + "\\resources\\avatar.png");

            wordMLPackage.save(new File(System.getProperty("user.dir") + "\\resources\\build\\testInsertImg_"+ConnectorLib.getDateTimeStr()+".docx"));

            ConnectorLib.trace("Test testInsertImg : " + true);
        }catch (Exception ex) {
            ConnectorLib.trace("testInsertImg - Exception : " + ex);
        }
    }

    public static void testFindContent(){
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "\\resources\\patern-signets.docx"));

            ConnectorLib.findContent(wordMLPackage, "CodePostal");

            ConnectorLib.trace("Test testFindContent : " + true);
        }catch (Exception ex) {
            ConnectorLib.trace("testFindContent - Exception : " + ex);
        }
    }
}