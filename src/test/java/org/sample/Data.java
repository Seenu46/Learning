package org.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data {
	public static void main(String[] args) throws IOException {
		
		File file=new File("C:\\Users\\Seenu\\eclipse-workspace\\ABC\\Excel\\Samp.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook workbook=new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		
		for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				CellType type = cell.getCellType();
			
			switch (type) {
			case STRING:
				String string = cell.getStringCellValue();
				System.out.println(string);
				break;
								
				
			case NUMERIC:
				if(DateUtil.isCellDateFormatted(cell)) {
					Date value = cell.getDateCellValue();
					SimpleDateFormat dateformat=new SimpleDateFormat("dd/MM/yy");
					String format = dateformat.format(value);
					System.out.println(format);
				}	
			else {
			}
				double d = cell.getNumericCellValue();
				BigDecimal b=new BigDecimal(d);
				String a = b.toString();
				System.out.println(a);
				
				break;
				
			default:
				
				break;
			}
			
			}
		}
		
	}
}

