/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.epitech.oliver_f.astextexls;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.net.URI;
import java.nio.file.FileVisitOption;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author florianoliver
 */
public class ReadXLSFiles {

    private String pathToFolder = null;

    public ReadXLSFiles(String pathToFolder) {
        this.pathToFolder = pathToFolder;
    }

    public ReadXLSFiles() {
    }

    public List<ResultRow> parse() {
        List<ResultRow> resultList = null;
        File folder = new File(pathToFolder);
        if (!folder.isDirectory()) {
            System.err.println("The specified folder is not a folder");
            return null;
        }
        List<Path> lpath = null;
        try {
            lpath = Files.list(Paths.get(pathToFolder)).filter(p -> p.toString().toLowerCase().endsWith(".xlsx")).collect(Collectors.toList());
        } catch (IOException ex) {
            Logger.getLogger(ReadXLSFiles.class.getName()).log(Level.SEVERE, null, ex);
        }
        if (lpath != null) {
            System.out.println("size " + lpath.size());
            resultList = parseAllFiles(lpath);
        }
        return resultList;
    }

    private List<ResultRow> parseAllFiles(List<Path> paths) {
        List<ResultRow> resultList = new ArrayList<ResultRow>();
        for (Path path : paths) {
            try {
                System.out.println("file : " + path.toAbsolutePath());
                FileInputStream file = new FileInputStream(path.toFile());
                Workbook wb = WorkbookFactory.create(file);
                Sheet sheet = wb.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                boolean found = false;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    //For each row, iterate through all the columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        //Check the cell type and format accordingly
                        String res = null;
                        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            double inte = cell.getNumericCellValue();
                            res = Double.toString(inte);
                        }
                        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            res = cell.getStringCellValue();
                        }
                        if (res != null && res.trim().toLowerCase().equals("login \névaluateur")) {
                            found = true;
                        }
                    }
                    if (found) {
                        ResultRow rr = new ResultRow();
                        Row rowFound = rowIterator.next();
                        Iterator<Cell> c = rowFound.cellIterator();
                        while (c.hasNext()) {
                            Cell cel = c.next();
                            String res = null;
                            if (cel.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                double inte = cel.getNumericCellValue();
                                res = Double.toString(inte);
                            }
                            if (cel.getCellType() == Cell.CELL_TYPE_STRING) {
                                res = cel.getStringCellValue();
                            }
                            rr.result.add(res);
                        }
                        resultList.add(rr);
                        found = false;
                        break;
                    }
                }
                file.close();
            } catch (IOException | InvalidFormatException e) {
            }
        }
        return resultList;
    }

    /**
     * @return the pathToFolder
     */
    public String getPathToFolder() {
        return pathToFolder;
    }

    /**
     * @param pathToFolder the pathToFolder to set
     */
    public void setPathToFolder(String pathToFolder) {
        this.pathToFolder = pathToFolder;
    }

}
