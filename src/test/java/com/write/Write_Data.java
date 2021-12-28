package com.write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {
	public static void main(String[] args) throws IOException {
		
		System.out.println("create a excel sheet");
		
		File f = new File("C:\\Users\\S.R\\eclipse-workspace\\Data_Driven\\STUDENTS DATA.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		//CREATE SHEET
		Sheet createSheet = wb.createSheet("STUD");
		//CREATE ROW
		Row createRow = createSheet.createRow(0);
		//CREATE CELL
		Cell createCell = createRow.createCell(0);
		//SET THE VALUE
		createCell.setCellValue("user data");
		//SET THE VALUE IN THE SECOND CELL
		wb.getSheet("STUD").getRow(0).createCell(1).setCellValue("password");
		
		wb.getSheet("STUD").createRow(1).createCell(0).setCellValue("mani");
		
		wb.getSheet("STUD").getRow(1).createCell(1).setCellValue("12345");
		
		
		FileOutputStream fos = new FileOutputStream(f);
		//WRITE
		wb.write(fos);
		//CLOSE
		wb.close();
		
		System.out.println("excel sheet created");
		
		
		
		
		
		
		
		
		
	}

}
