package Analyzer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;



public class ExcelAnalyzer {

    public static void main(String[] args) {
        try {
            FileInputStream excelFile = new FileInputStream("D:\\Talathi\\Assignment_Timecard.xls");
            Workbook workbook = new HSSFWorkbook(excelFile); 
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet
            
           
            int consecutiveCount = 0;
            String pid="";
            String pdid="";
            for (Row row : sheet) {
            	String id = row.getCell(0).getStringCellValue();
                consecutiveCount = 0;
                
              
                for (int cellIndex = 2; cellIndex < row.getLastCellNum(); cellIndex++) {
                	
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null && pid.equals(id)) {
                    	
                    	consecutiveCount++;
                    }
                    if (consecutiveCount >= 7 && pdid!=id ) {
                        System.out.println("Employee: " + id + ", has worked for 7 consecutive days" );
                        pdid=id;
                    }
                    pid=id;
                }
                
                
            }

            excelFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

