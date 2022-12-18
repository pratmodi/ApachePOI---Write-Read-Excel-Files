package com.pratmodi.one;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) {
		Test t = new Test();
	//	t.readAllRows();
		t.deleteRow("user123");
	}

	public void readAllRows() {
		try {
			FileInputStream file = new FileInputStream(new File("C:\\Java Practice Workspace\\test.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					case NUMERIC:
						System.out.print((int) cell.getNumericCellValue() + "\t");
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					default:
						break;
					}
				}
				System.out.println("");
			}
			file.close();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void deleteRow(String username) {
		try {
			FileInputStream file = new FileInputStream(new File("C:\\Java Practice Workspace\\test.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell;
			Row row = null;
			Row temp;
			int serial = 1;
			Iterator<Row> rowIterator = sheet.iterator();
			int row_count = 0;
			List<Row> toRemove = new ArrayList<Row>();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
//				Iterator<Cell> cellIterator = row.cellIterator();

				if (username.trim().equals(row.getCell(1).getStringCellValue())) {
					System.out.println("***********ONE*************");
					System.out.println(row.getCell(1).getStringCellValue());
					toRemove.add(row);
				}
				row_count++;
			}
			
			// loop the list and call sheet.removeRow() on every entry
			for (Row r : toRemove) {
				sheet.removeRow(r);
				sheet.shiftRows(r.getRowNum(), sheet.getLastRowNum(), 1);
			}

//			for (int i = 1; i < sheet.getLastRowNum(); i++) {
//				if ((temp = sheet.getRow(i)) != null) {
//					cell = temp.getCell(0);
//					cell.setCellValue(serial++);
//				}
//				
//			}
			
			FileOutputStream out = new FileOutputStream(new File("C:\\Java Practice Workspace\\test.xlsx"));
            workbook.write(out);
            out.close();

		} catch (Exception e) {

		}
	}

}
