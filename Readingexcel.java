 package com.excelconcept;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readingexcel {

	public static void main(String[] args) {
		
		//path of an excel file
		String filepath="Test.xlsx";
		
		try {
			//create input stream of fileinputstream object
			FileInputStream inpstream=new FileInputStream(filepath);
			
			//create on object of workbook
			XSSFWorkbook workbook=new XSSFWorkbook(inpstream);
			
			//get an access to sheet
			Sheet sheet=workbook.getSheetAt(0);
			
			//iterate over row in sheet
			for(Row row:sheet) {
				//iterate over the cells of each row
				for(Cell cell:row) {
					//check if the data is string or number
					if(cell.getCellType()==CellType.STRING) {
						System.out.print(cell.getStringCellValue()+"\t");
					}else if(cell.getCellType()==CellType.NUMERIC) {
						System.out.print(cell.getNumericCellValue()+"\t");
					}
				}
				System.out.println();
			}
			
		}catch(IOException e) {
			e.printStackTrace();
		}
		
	}

}
