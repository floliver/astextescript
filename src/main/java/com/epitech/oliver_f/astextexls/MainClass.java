/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.epitech.oliver_f.astextexls;

import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

/**
 *
 * @author florianoliver
 */
public class MainClass {

    private static final Options options = new Options();

    private static void initalizeOption() {
        options.addOption("folder", true, "folder of all xlsx files from the students");
        options.addOption("file", true, "file to fill");
        options.addOption("help", false, "print this help");
    }

    public static void main(String[] args) {
        initalizeOption();
        CommandLineParser parser = new BasicParser();
        CommandLine cl = null;
        try {
            cl = parser.parse(options, args);
        } catch (ParseException ex) {
            Logger.getLogger(MainClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        if (cl != null && cl.hasOption("help")) {
            HelpFormatter hf = new HelpFormatter();
            hf.printHelp("astexte script", options);
        }
//        System.out.println("OK");
//        ReadXLSFiles rxlsfile = new ReadXLSFiles("../rendu");
//        List<ResultRow> lRows = rxlsfile.parse();
//        if (lRows != null) {
//            System.out.println("bien ouej");
//            WriteXLSFile writeXLSFile = new WriteXLSFile(lRows, "../B1-EE-LETTRE-ARGUMENTEE-fichier-notes_ALL-MODELE.xlsx");
//            writeXLSFile.write();
//        }
        if (cl.hasOption("folder") && cl.hasOption("file")) {
            System.out.println("OK");
            ReadXLSFiles readXLSFiles = new ReadXLSFiles(cl.getOptionValue("folder"));
            List<ResultRow> lRows = readXLSFiles.parse();
            System.out.println("reading ..." + lRows.size());
            lRows.stream().forEach(p -> System.out.println("p : " + p));
            WriteXLSFile writeXLSFile = new WriteXLSFile(lRows, cl.getOptionValue("file"));
            writeXLSFile.write();

        } else {
            HelpFormatter hf = new HelpFormatter();
            hf.printHelp("astexte script", options);
        }

    }

}
