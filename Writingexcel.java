package com.excelconcept;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writingexcel {

	public static void main(String[] args) {
		
		//create a blank excel sheet of XSSFWorkbook
		//try with resource				
        try(XSSFWorkbook workbook=new XSSFWorkbook()){
        	
        //create sheet	
        	
        XSSFSheet sheet=workbook.createSheet("Sheet1");
       
        //create an array of objects
        Object[][]data= {
        		{"name","age","city"},
        		{"lavish",39,"pune"},
        		{"swapnil",37,"mumbai"},
        		{"himanshu",36,"delhi"},
        		{"pavan",45,"hyderabad"},
        		
        };
        
        //writing the data into excel
        int rowNum=0;
        for(Object[] rowdata:data) {
        	//create a row in the sheet
        	XSSFRow row=sheet .createRow(rowNum++);
        	
        	//insert the data into cells
        	int colNum=0; //to print the column number
        	for(Object field:rowdata) {
        		Cell cell=row.createCell(colNum++); //created cell
        		if(field instanceof String) {
        			cell.setCellValue((String)field);        			
        		}else if(field instanceof Integer) {
        			cell.setCellValue((Integer)field);
        		}
        			
        			
        	}
        }
        	//create file outstream object and write the data
            try(FileOutputStream os=new FileOutputStream("test.xlsx")){
            	workbook.write(os);
            }
            System.out.println("Data added successfully to file...");
            
        }catch (IOException e) {
        	e.printStackTrace();
        }
        
        
	}

}
