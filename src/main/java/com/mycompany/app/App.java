package com.mycompany.app;

/**
 * Hello world!
 *
 */
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

		// TODO Auto-generated method stub
		public static XSSFSheet ExcelWSheet;
		public static XSSFWorkbook ExcelWorkbook;
		public XSSFCell ExcelCell;
		public static XSSFRow ExcelRow;
		public static String filename_path="/Employee_Data.xlsx";
		
		public static void main(String args[]) throws IOException
		{
                 InputStream url = App.class.getClass().getResourceAsStream("/Employee_Data.xlsx");

		FileInputStream ExcelFile = new FileInputStream(App.class.getResourceAsStream("/Employee_Data.xlsx"));
		ExcelWorkbook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWorkbook.getSheet("Sheet1");
		//int rowCount= ExcelWSheet.getLastRowNum() -ExcelWSheet.getFirstRowNum();
		int rowCount= 3;
		
		 for (int i = 0; i < rowCount+1; i++) {

			 ExcelRow = ExcelWSheet.getRow(i);

		        //Create a loop to print cell values in a row

		        for (int j = 0; j < ExcelRow.getLastCellNum(); j++) {
	

		            //Print Excel data in console

		            System.out.print(ExcelRow.getCell(j).getStringCellValue()+" | ");

		        }

		        System.out.println();
		 }
			
		}

	}

