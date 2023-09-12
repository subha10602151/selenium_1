package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class sample {
public static void main (String[] args) throws Exception
{
	File f= new File("C:\\Users\\Gopisan\\Desktop\\demo.xlsx");
	    FileOutputStream f1 = new FileOutputStream(f);
		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet sheet= workbook.createSheet("IPL");
		XSSFRow row= sheet.createRow(0);
		XSSFCell cell1=row.createCell(0);
		XSSFCell cell2=row.createCell(1);
		cell1.setCellValue("MATCHES");
		cell2.setCellValue(123);
		workbook.write(f1);
		workbook.close();

        FileInputStream  f2 = new FileInputStream(f);
		XSSFSheet sheet2 = workbook.getSheet("IPL");
		XSSFRow row2 = sheet2.getRow(0);
		XSSFCell cell3 = row2.getCell(0);
		XSSFCell cell = row2.getCell(1);
		String text = cell3.getStringCellValue();
		System.out.println(text);
		double num = cell.getNumericCellValue();
		System.out.println(num);
		workbook.close();
		
}
}
