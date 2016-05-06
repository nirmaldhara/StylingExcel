# StylingExcel
##Styling Excel Sheets using java POI

Apache POI is the api to create and modify Microsoft office files.
But here i will discuss about how to write data to excel sheets(spreadsheet) with Styling.
sometimes we need to color and bold the text, even  backgrund color also.
Below code can change the text style, color, backgroud , border. you can do many more using XSSFCellStyle  class.

Below jar files are require to write excel sheets with Styling .

1. ooxml-schemas-1.3.jar
2. poi-3.14.jar
3. poi-ooxml-3.14.jar
4. xmlbeans-2.6.0.jar

#Code

```java

package com.javaant;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
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
public class ReadWriteExcel {
	static String filepath = null;

	public static void main(String ar[]) throws IOException {
		ReadWriteExcel rw = new ReadWriteExcel("D:\\Java_Ant_Post\\StylingExcel\\excels\\abc.xlsx");
		rw.writeDataToExcel(filepath);

	}

	public ReadWriteExcel(String filepath) {
		ReadWriteExcel.filepath = filepath;
	}

	public File getFile() throws FileNotFoundException {
		File here = new File(filepath);
		return new File(here.getAbsolutePath());

	}

	private static void writeToCell(int rowno, int colno, XSSFSheet sheet, XSSFCellStyle myStyle, String val) {
		try {
			sheet.getRow(rowno);
			XSSFRow row = sheet.getRow(rowno);
			if (row == null) {
				row = sheet.createRow(rowno);
			}
			XSSFCell cell = row.createCell(colno);
			cell.setCellStyle(myStyle);
			cell.setCellValue(val);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static XSSFCellStyle cellStyle(XSSFWorkbook wb, String fontStyle, String backGroundColor, String color) {
		XSSFCellStyle myStyle = wb.createCellStyle();
		if (fontStyle.equalsIgnoreCase("yes")) {
			XSSFFont font = wb.createFont();
			font.setFontHeightInPoints((short) 16);
			font.setColor(IndexedColors.WHITE.getIndex());
			font.setBold(true);

			myStyle.setFont(font);
		}
		if (backGroundColor.equalsIgnoreCase("yes")) {
			if (color.equalsIgnoreCase("green")) {
				myStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
			} else if (color.equalsIgnoreCase("red")) {
				myStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			}

			myStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}

		myStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		myStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		myStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		myStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);

		return myStyle;
	}

	public void writeDataToExcel(String file) throws IOException {
		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		FileOutputStream fileOut = null;

		String excelFileName = file;

		String sheetName = "Sheet1";

		wb = new XSSFWorkbook();
		sheet = wb.createSheet(sheetName);
		writeToCell(0, 0, sheet, cellStyle(wb, "yes", "yes", "green"), "Wel come to JavaAnt");
		writeToCell(1, 0, sheet, cellStyle(wb, "yes", "yes", "green"), "Get Logic & code ");
		writeToCell(0, 1, sheet, cellStyle(wb, "yes", "yes", "green"), "Date- " + new Date().toString());
		writeToCell(1, 1, sheet, cellStyle(wb, "yes", "yes", "red"), "Plese share this site");
		writeToCell(0, 2, sheet, cellStyle(wb, "yes", "yes", "green"), "Hepl all to solve the problem ");
		writeToCell(1, 2, sheet, cellStyle(wb, "yes", "yes", "red"), "all java technologyies  ");
		writeToCell(3, 1, sheet, cellStyle(wb, "yes", "yes", "green"), "Core java");
		writeToCell(3, 0, sheet, cellStyle(wb, "yes", "yes", "green"), "Jsp");
		writeToCell(3, 2, sheet, cellStyle(wb, "yes", "yes", "green"), "Servlets");
		int r = 4;

		System.out.println("working fine");
		fileOut = new FileOutputStream(excelFileName);
		wb.write(fileOut);

	}

}
```
