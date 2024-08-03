package com.excelconcept;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Countingexcel {

	public static void main(String[] args) {
		
		//provide path of file
		String filepath="Test.xlsx";
		
		try {
			FileInputStream inpstream=new FileInputStream(filepath);
			XSSFWorkbook workbook=new XSSFWorkbook(inpstream);
			Sheet sheet=workbook.getSheetAt(0);
			
			int lastrownum=sheet.getLastRowNum();
			int firstrownum=sheet.getFirstRowNum();
			
			System.out.println(firstrownum);
			System.out.println(lastrownum);
			
			int rowcount=sheet.getLastRowNum()-sheet.getFirstRowNum();
			
			System.out.println("Total Rows "+rowcount);
			
		}catch (Exception e) {
		   e.printStackTrace();	
		}

	}

}
