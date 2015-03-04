/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.epitech.oliver_f.astextexls;

import java.io.File;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.scene.Scene;
import javafx.scene.layout.StackPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
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
public class MainClass extends Application {

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
        if (cl.hasOption("folder") && cl.hasOption("file")) {
            launchWriteAndRead(cl.getOptionValue("folder"), cl.getOptionValue("file"));
            System.out.println("OK");

        } else {
            HelpFormatter hf = new HelpFormatter();
            hf.printHelp("astexte script", options);
        }
        launch(args);
    }

    private static void launchWriteAndRead(String folder, String file) {
        ReadXLSFiles readXLSFiles = new ReadXLSFiles(folder);
        List<ResultRow> lRows = readXLSFiles.parse();
        System.out.println("reading ..." + lRows.size());
        lRows.stream().forEach(p -> System.out.println("p : " + p));
        WriteXLSFile writeXLSFile = new WriteXLSFile(lRows, file);
        writeXLSFile.write();
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        primaryStage.setTitle("Hello World!");
        DirectoryChooser dChooser = new DirectoryChooser();
        dChooser.setTitle("Choose the directory of the excel files");
        File defaultDirectory = new File("c:/");
        dChooser.setInitialDirectory(defaultDirectory);
        File selectedDirectory = dChooser.showDialog(primaryStage);
        StackPane root = new StackPane();
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Open excel 'ALL' File");
        File choosenFile = fileChooser.showOpenDialog(primaryStage);
        launchWriteAndRead(selectedDirectory.getAbsolutePath(),  choosenFile.getAbsolutePath());
        //root.getChildren().add(btn);
        //primaryStage.setScene(new Scene(root, 300, 250));
        //primaryStage.show();
        Platform.exit();

    }

}
