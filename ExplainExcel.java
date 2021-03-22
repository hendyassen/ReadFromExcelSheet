package com.crm.testcases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExplainExcel {
	/*public static void main(String[] args) throws Exception {
		File file = new File("C:\\Users\\k.mina\\Desktop\\DataTest.xlsx");
		// to get the eexcel sheet
		FileInputStream fis = new FileInputStream(file);

		// Apache POI
		// declare the workbook of the excelsheet
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		/*
		 * there are two types of excel sheet 1- .xls 2- .xlsx
		 
		// the declare the sheet
		// we can get sheet by name or index

		// this by index
		Sheet sheet = workbook.getSheetAt(0);
		// this is by name
		// Sheet sheet = workbook.getSheet("Sheet1");
		String FamilyName = sheet.getRow(1).getCell(1).toString();
		System.out.println(FamilyName);

		// to close
		workbook.close();
		
		//to get the value of the last row
		//same to the colum
		int rows = sheet.getLastRowNum();
		int columns = sheet.getRow(0).getLastCellNum();
		Object data[][] = new Object[rows][columns];
		for(int i=0; i<rows; i++) {
			for(int k=0; k<columns; k++) {
				data[i][k] = sheet.getRow(i).getCell(k);
			}
		} 

	} */
	

}
