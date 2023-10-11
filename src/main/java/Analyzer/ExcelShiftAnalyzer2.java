package Analyzer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Time;

public class ExcelShiftAnalyzer2 {

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
                    
                        if (hours > 14.0) {
                        	if (pid!=id) {
                        		pid=id;
                            System.out.println("Employee: " + id + " time "+name+", has worked for more than 14 hours in a single shift.");
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

