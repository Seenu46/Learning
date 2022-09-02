package org.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewCell {
	 public static void main(String[] args) throws Exception {
		File file= new File("C:\\Users\\Seenu\\eclipse-workspace\\ABC\\Excel\\neww.xlsx");
		//FileOutputStream stream =new FileOutputStream(file);
		Workbook workbook = new  XSSFWorkbook();
		
		Sheet sheet = workbook.createSheet("Sheet2");
		Row createRow = sheet.createRow(0);
		Cell createCell = createRow.createCell(3);
		createCell.setCellValue("Seenivasan");
		
		FileOutputStream stream1=new FileOutputStream(file);
		workbook.write(stream1);
		 
	}

}
