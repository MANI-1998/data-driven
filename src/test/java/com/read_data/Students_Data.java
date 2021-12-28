package com.read_data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Students_Data {
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\S.R\\eclipse-workspace\\Data_Driven\\STUDENTS DATA.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheet("students");
		int row_size = sheet.getPhysicalNumberOfRows();
		
		for (int i = 0; i < row_size; i++) {
			Row row = sheet.getRow(i);
			int cell_size = row.getPhysicalNumberOfCells();
			
			for (int j = 0; j < cell_size; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				
				if (cellType.equals(cellType.STRING)) {
					String svalue = cell.getStringCellValue();
					System.out.println(svalue);
				}
				else if (cellType.equals(cellType.NUMERIC)) {
					double nvalue = cell.getNumericCellValue();
					int value = (int) nvalue ;
					System.out.println(value);
				}
				
			}
			System.out.println("***********");
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
