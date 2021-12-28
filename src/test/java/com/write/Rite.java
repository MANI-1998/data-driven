package com.write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Rite {
	public static void main(String[] args) throws IOException {
		
		System.out.println("excel sheet created");
		
		File f = new File("C:\\Users\\S.R\\eclipse-workspace\\Data_Driven\\STUDENTS DATA.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet cs = wb.createSheet("java");
		
		Row cr0 = cs.createRow(0);
		
		cr0.createCell(0).setCellValue("mani");
		
	    cr0.createCell(1).setCellValue("java");
		
		Row cr1 = cs.createRow(1);
		
		cr1.createCell(0).setCellValue("tina");
		
		cr1.createCell(1).setCellValue("python");
		
		Row cr2 = cs.createRow(2);
		
		cr2.createCell(0).setCellValue("abi");
		
		cr2.createCell(1).setCellValue("spacex ceo");
		
		FileOutputStream fos = new FileOutputStream(f);
		
		wb.write(fos);
		
		wb.close();
		
		System.out.println("data written successfully");
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
