package com.cloudera.excel;

//import java.io.File;
//import java.io.FileInputStream;
import java.io.InputStream;
//import java.util.ArrayList;
import java.util.Iterator;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToObject {
	
	private StringBuilder currentString = null;
	private long bytesRead = 0;

//	public static void main(String[] args)
//	{
//		new ExcelToObject().ExcelToObjectConversion();
//	}
	
	public String ExcelToObjectConversion(InputStream inpstrm) 
	{
        
		try
         {
             //FileInputStream inpstrm = new FileInputStream(new File("/home/cloudera/Sumit/Datasets/Name.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(inpstrm);

			// Taking first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			currentString = new StringBuilder();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						bytesRead++;
						currentString.append(cell.getBooleanCellValue() + "\t");
						//System.out.println(cell.getBooleanCellValue() + "\t");
						break;

					case Cell.CELL_TYPE_NUMERIC:
						bytesRead++;
						currentString.append(cell.getNumericCellValue() + "\t");
						//System.out.println(cell.getNumericCellValue() + "\t");
						break;

					case Cell.CELL_TYPE_STRING:
						bytesRead++;
						currentString.append(cell.getStringCellValue() + "\t");
						//System.out.println(cell.getStringCellValue() + "\t");
						break;	
					default:
						if(HSSFDateUtil.isCellDateFormatted(row.getCell(0)))
						{
							bytesRead++;
							currentString.append(cell.getDateCellValue() + "\t");
						}
					}
				}
				currentString.append("\n");
				//System.out.println("\n");
			}
			inpstrm.close();
         } 
         catch (Exception e) 
         {
             e.printStackTrace();
         }
		return currentString.toString();
     }
	
	public long getBytesRead() {
		return bytesRead;
	}
}
