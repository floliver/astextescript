/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.epitech.oliver_f.astextexls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
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
public class WriteXLSFile {

    private final List<ResultRow> results;
    private final String pathToFile;

    public WriteXLSFile(List<ResultRow> results, String path) {
        this.results = results;
        this.pathToFile = path;
    }

    public void write() {
        FileInputStream file = null;
        try {
            file = new FileInputStream(pathToFile);
            Workbook wb = WorkbookFactory.create(file);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            int i = 0;
            int listIndex = 0;
            while (rowIterator.hasNext() && listIndex < results.size()) {
                Row row = rowIterator.next();
                if (i > 1) {
                    Iterator<Cell> cellIterator = row.cellIterator();;
                    int cellIndex = 0;
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String r = results.get(listIndex).result.get(cellIndex);
                        try {
                            if (r == null)
                                throw new NumberFormatException();
                            Double resDouble = Double.parseDouble(r);
                            Integer resInt = resDouble.intValue();
                            cell.setCellValue(resInt.toString());
                        } catch (NumberFormatException e) {
                            cell.setCellValue(results.get(listIndex).result.get(cellIndex));
                        }
                        cellIndex++;
                    }
                    listIndex++;
                }
                i++;
            }
            System.out.println("listindex " + listIndex);
            file.close();
            FileOutputStream outFile = new FileOutputStream(new File(pathToFile));
            wb.write(outFile);
            outFile.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(WriteXLSFile.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(WriteXLSFile.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(WriteXLSFile.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                file.close();
            } catch (IOException ex) {
                Logger.getLogger(WriteXLSFile.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

}
