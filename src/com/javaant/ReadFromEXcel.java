package com.javaant;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromEXcel {

	static String filepath = null;
	Object[][] data = null;
	public static void main(String ar[]) {
		ReadFromEXcel rw = new ReadFromEXcel("D:\\Java_Ant_Post\\readWritexls-poi\\ReadWriteExcel\\excels\\abc.xlsx");
		rw.writeDataToExcel(filepath);
		//rw.readDataFromExcel();

	}
	public ReadFromEXcel(String filepath) {
		this.filepath = filepath;
	}

	public File getFile() throws FileNotFoundException {
		File here = new File(filepath);
		return new File(here.getAbsolutePath());

	}

	
	
	private static void writeToCell(int rowno, int colno, XSSFSheet sheet, String val) {
		try {
			sheet.getRow(rowno);
			XSSFRow row = sheet.getRow(rowno);
			if (row == null) {
				row = sheet.createRow(rowno);
			}
			XSSFCell cell = row.createCell(colno);
			cell.setCellValue(val);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
	public static void writeDataToExcel(String file) {
		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		FileOutputStream fileOut = null;
		
		
			String excelFileName = file;

			String sheetName = "Sheet1";//name of sheet

			wb = new XSSFWorkbook();
			sheet = wb.createSheet(sheetName);
			DecimalFormat df2 = new DecimalFormat(".##");
			try{

	
			
			writeToCell(0, 0, sheet,  "This is one");// row 1 column 1
			writeToCell(0, 1, sheet,  "this is two");// row 1 column 2
			writeToCell(1, 0, sheet,  "this is three");// row 2 column 1
			writeToCell(1, 1, sheet,  "this is four");// row 2 column 2
			int r = 4;
			
			System.out.println("working fine");
			fileOut = new FileOutputStream(excelFileName);
			wb.write(fileOut);

			//write this workbook to an Outputstream.
		} catch (Exception e) {
			e.printStackTrace();
		} finally {

			try {
				fileOut.flush();
				fileOut.close();
			} catch (IOException ex) {
				Logger.getLogger(ReadFromEXcel.class.getName()).log(Level.SEVERE, null, ex);
			}

		}
	}

	
}
