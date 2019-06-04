package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	private static XSSFWorkbook excelBook;  //
	private static XSSFSheet excelSheet;
	private static XSSFRow excelRow;
	private static XSSFCell excelCell;
	private static String excelFilePath;
	
	/**
	 * This method accepts path of Excel and Sheet name and opens Excel file
	 * @param path
	 * @param sheetName
	 */
	
	public static void openExcelFile(String path, String sheetName) {
		
		excelFilePath = path;  //we this only if names are same
		
		
		try {
			File file = new File(path);
			FileInputStream input = new FileInputStream(file);
			excelBook = new XSSFWorkbook(input);
			excelSheet = excelBook.getSheet(sheetName);
			
			
		} catch (Exception e) {
			System.out.println("Excel file not found");
		}
	}
	
	/**
	 * This method accepts row number and cell number and returns value of the given Excel cell
	 * 
	 * @param row
	 * @param cell
	 * @return Value of Cell
	 */
	
	public static String getCellValue(int row, int cell) {
		
		String value = excelSheet.getRow(row).getCell(cell).toString();
		
		return value;
	}
	
	/**
	 * This method accepts new value of cell, row number and cell number and changes the value of the 
	 * given Excel cell
	 * 
	 * @param value
	 * @param row
	 * @param cell
	 */
	public static void setCellValue(String value, int row, int cell) {
		
		excelSheet.getRow(row).getCell(cell).setCellValue(value);
		
	}
	
	/**
	 * This methods returns all number of rows in our Excel Sheet
	 * 
	 * @return Number of Rows
	 */
	
	public static int getNumberOfRows() {
		
		return excelSheet.getPhysicalNumberOfRows();
	}
	
	

}
