package com.read_data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Data {

	public static void main(String[] args) throws IOException {

           //to read a file
		File f = new File("C:\\Users\\S.R\\eclipse-workspace\\Data_Driven\\STUDENTS DATA.xlsx");
           //to read a file data values
		FileInputStream fis = new FileInputStream(f);
           //to read a excel sheet
		Workbook wb = new XSSFWorkbook(fis);
		//to get the sheet
		Sheet sheetAt = wb.getSheetAt(0);
		//to get the number of rows in sheet
		int row_size = sheetAt.getPhysicalNumberOfRows();
		
		for (int i = 0; i < row_size; i++) {
			//to get into the row
			Row row = sheetAt.getRow(i);
			//to get the number of cells in the row
			//so to get the cells in every row we use for loop
			int cell_size = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cell_size; j++) {
				
			
			
				//to get the cell from the row
				Cell cell = row.getCell(0);
				//to get the cell from the celltype
				CellType cellType = cell.getCellType();
			
				if (cellType.equals(cellType.STRING)) {
					String svalue = cell.getStringCellValue();
					System.out.println("  "+svalue);
				}
				else if (cellType.equals(cellType.NUMERIC)) {
					double nvalue = cell.getNumericCellValue();
					int value = (int) nvalue;
					System.out.println("  "+value);
					
				}
				
				
			}
			
			}
			
			
		}

	}


