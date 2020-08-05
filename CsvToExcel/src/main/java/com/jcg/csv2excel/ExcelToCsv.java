package com.jcg.csv2excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.opencsv.CSVWriter;


public class ExcelToCsv {

    public static void echoAsCSV(Sheet sheet) throws IOException {
    	//Instantiating the CSVWriter class
        CSVWriter writer = new CSVWriter(new FileWriter("C:\\Users\\AndyY\\git\\CsvExcel\\CsvToExcel\\config\\output.csv"));
        //Writing data to a csv file
        
    	
    	Row row = null;
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
            	String line[] = {"\"" + row.getCell(j) + "\","};
            	writer.writeNext(line);
            	
            	System.out.print("\"" + row.getCell(j) + "\",");
            }
            System.out.println();
        }
        writer.close();
    }

   
    public static void main(String[] args) {
        InputStream inp = null;
        try {
            inp = new FileInputStream("C:\\Users\\AndyY\\git\\CsvExcel\\CsvToExcel\\config\\EXCEL_DATA.xlsx");
            Workbook wb = WorkbookFactory.create(inp);

            for(int i=0;i<wb.getNumberOfSheets();i++) {
                System.out.println(wb.getSheetAt(i).getSheetName());
                echoAsCSV(wb.getSheetAt(i));
            }
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ExcelToCsv.class.getName()).log(Level.SEVERE, null, ex);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelToCsv.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelToCsv.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                inp.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelToCsv.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
}

