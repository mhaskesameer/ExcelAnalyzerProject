package Analyzer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelShiftAnalyzer {

    public static void main(String[] args) {
        try {
            FileInputStream excelFile = new FileInputStream("D:\\Talathi\\Assignment_Timecard.xls");
            Workbook workbook = new HSSFWorkbook(excelFile); 
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet
            String pid="";
            for (Row row : sheet) {
            	String id = row.getCell(0).getStringCellValue();
                String name = row.getCell(4).getStringCellValue();
                
              
                for (int cellIndex = 2; cellIndex < row.getLastCellNum(); cellIndex++) {
                	
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    	String[] timeSplit = name.split(":");

                        //convert arrays to doubles 
                        double hours = Double.parseDouble(timeSplit[0]);
                    
                       
                        if (hours > 1.0 && hours<10.0) {
                        	if (pid!=id) {
                        		pid=id;
                                System.out.println("Employee: " + id + ", Time: " + name + " has consecutive shifts between 1 and 10 hours apart.");
							}
                        	
                        }
                    }
                }

                
            }

            excelFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

