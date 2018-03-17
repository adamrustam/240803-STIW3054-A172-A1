import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.DataFormatter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.Writer;

import java.util.Iterator;/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author adamrustam
 */
public class ApachePOIExcelRead {
    
    
    private static final String FILE_NAME = "/Users/adamrustam/Desktop/sample.xlsx";

    public static void main(String[] args) {
        
        Writer writer = null;
        boolean line = true;
        
        try {
            
            DataFormatter df = new DataFormatter();

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            
            File file = new File("/Users/adamrustam/240803-STIW3054-A172-A1.wiki/Home.md");
            writer = new BufferedWriter(new FileWriter(file));
               
            
          

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String val = df.formatCellValue(currentCell);
                    
                    System.out.print(val +"|");
                  
                    
                    writer.write(val +"|"); 

                }
                System.out.println();
                writer.write("\n");
                if (line==true){
                    writer.write("---|---|---|---|\n");
                    line=false;
                }
                

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        try {                 
            if (writer != null) {                     
                writer.close();                 
            }             
        } catch (IOException e) {                 
            e.printStackTrace();             
        }         
        
    }

}




