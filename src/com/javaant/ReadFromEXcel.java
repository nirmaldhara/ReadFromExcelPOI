package com.javaant;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromEXcel {


	static String filepath = null;
	static Object[][] data = null;

	public static void main(String ar[]) {
		ReadFromEXcel rw = new ReadFromEXcel("D:\\Java_Ant_Post\\readWritexls-poi\\ReadWriteExcel\\excels\\abc.xlsx");
		data = rw.readDataFromExcel();
		// rw.readDataFromExcel();

	}

	public ReadFromEXcel(String filepath) {
		this.filepath = filepath;
	}

	public Object[][] readDataFromExcel() {
		final DataFormatter df = new DataFormatter();
		try {

			FileInputStream file = new FileInputStream(getFile());
			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			int rownum = 0;
			int colnum = 0;
			// ignore the first row , it may be header
			Row r = rowIterator.next();

			int rowcount = sheet.getLastRowNum();
			int colcount = r.getPhysicalNumberOfCells();
			data = new Object[rowcount][colcount];

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				colnum = 0;
				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					data[rownum][colnum] = df.formatCellValue(cell);
					System.out.print(" " + df.formatCellValue(cell));
					colnum++;
				}
				rownum++;
				System.out.println("");
			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		return data;
	}

	public File getFile() throws FileNotFoundException {
		File here = new File(filepath);
		return new File(here.getAbsolutePath());

	}

	
}
