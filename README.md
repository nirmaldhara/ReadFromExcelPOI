# Read From Excel POI
Read data from excel using poi and java

Apache POI is the api to create and modify Microsoft office files. 
But here i will discuss about how to read data from excel sheets(spreadsheet)

Below jar files required to Create excel sheets.

##1. ooxml-schemas-1.3.jar

##2. poi-3.14.jar

##3. poi-ooxml-3.14.jar

##4. xmlbeans-2.6.0.jar

```java
package com.javaant;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author nirmal
 */
public class WriteToExcel {

	static String filepath = null;
	static Object[][] data = null;

	public static void main(String ar[]) {
		WriteToExcel rw = new WriteToExcel("D:\\Java_Ant_Post\\readWritexls-poi\\ReadWriteExcel\\excels\\abc.xlsx");
		data = rw.readDataFromExcel();
		// rw.readDataFromExcel();

	}

	public WriteToExcel(String filepath) {
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

```
