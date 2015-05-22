package com.bonitasoft;

import java.io.FileNotFoundException;
import java.nio.file.*;
import java.util.logging.Logger;

/**
 * Created by Fabrice.R on 20/05/2015.
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
}