package com.read_data;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Task {
	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\S.R\\eclipse-workspace\\Data_Driven\\STUDENTS DATA.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = wb.getSheet("students");
		int row_size = sheet.getPhysicalNumberOfRows();
		System.out.println("PARTICULAR STUDENT NAME");
		System.out.println("*********************************************************");
		for (int i = 0; i < 1; i++) {
			Row row = sheet.getRow(1);
			int cell_size = row.getPhysicalNumberOfCells();
			for (int j = 0; j < 1; j++) {
			Cell cell = row.getCell(0);
			CellType cellType = cell.getCellType();

			if (cellType.equals(cellType.STRING)) {
				String svalue = cell.getStringCellValue();
				System.out.print(" "+svalue);
			} else if (cellType.equals(cellType.NUMERIC)) {
				double nvalue = cell.getNumericCellValue();
				int value = (int) nvalue;
				System.out.print(" "+value);
			}
			}
		}
		
	}

}
