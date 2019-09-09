package com.varsha;
/**
 * @author Vaishali Varsha
 * @date Sept 08, 2019
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.Iterator;
import java.util.Map;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ReadXlsxTest {

	public static void main(String[] args) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {

		System.out.println("Run started");

		// Create an object of ReadXlsxTest class
		ReadXlsxTest objExcelFile = new ReadXlsxTest();

		// Prepare the path of excel file
		String filePath = System.getProperty("user.dir") + "\\excelExportAndFileIO\\";

		// Call read file method of the class to read data
		objExcelFile.readExcelFile(filePath, "allbills_Aug.xlsx", "");

		System.out.println("Run ended");
	}

	public void readExcelFile(String filePath, String fileName, String sheetName) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// Create an object of File class to open xlsx file
		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file
		FileInputStream inputStream = new FileInputStream(file);

//		Workbook myExcelBook = null;
		Workbook myExcelBook = WorkbookFactory.create(inputStream);
		// Find the file extension by splitting file name in substring and getting only
		// extension name
		/*String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file
		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class
//			myExcelBook = new XSSFWorkbook(inputStream);
			myExcelBook = WorkbookFactory.create(inputStream);

		} // Check condition if the file is xls file
		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of HSSFWorkbook class
//			myExcelBook = new HSSFWorkbook(inputStream);
			myExcelBook = WorkbookFactory.create(inputStream);
					

		}*/
		Sheet myExcelBookSheet;
		// Read sheet inside the workbook by its name
		if (sheetName != null && !sheetName.isEmpty()) {
			myExcelBookSheet = myExcelBook.getSheet(sheetName);
		} else {
			myExcelBookSheet = myExcelBook.getSheetAt(0);
		}

		// Finds the workbook instance for XLSX file
		// In case of an xls file format, the correct workbook
		// and sheet implementation to use is HSSF instead of
		// XSSF. The rest is the same.
//		XSSFWorkbook workBook = new XSSFWorkbook(fileInput);

		// Return first sheet from the XLSX workbook
//		XSSFSheet sheet = workBook.getSheetAt(0);

//		Using Iterator please expand further 
		// Get iterator to all the rows in current sheet
		Iterator<Row> rowIterator = myExcelBookSheet.iterator();

		// Traversing over each row of excel file
		// Moving 1 row forward to skip table headers
		Row row;
		while (rowIterator.hasNext()) {
			row = rowIterator.next();
			// For each row, iterate through each columns
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellTypeEnum()) {
	            case STRING: 
	            	//Print Excel STRING data in console
		            System.out.print(cell.getStringCellValue()+"|| "); 
		            break;
	            case NUMERIC: 
	            	//Print Excel data in console
	            	if (DateUtil.isCellDateFormatted(cell)) {
	            		//Print Excel DATE data in dateFormat to console
			            System.out.print(cell.getDateCellValue() + "|| "); 
                    } else {
                    	//Print Excel NUMERIC data to console
			            System.out.print(cell.getNumericCellValue() + "|| ");
                    }
	            	break;
	            case BOOLEAN: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print(cell.getBooleanCellValue() + "|| ");
		            break;
		        case FORMULA: 
		        	//Print Excel BOOLEAN data to console
		        	switch(cell.getCachedFormulaResultTypeEnum()) {
		            case NUMERIC:
		            	System.out.print("Last evaluated as: " + cell.getNumericCellValue() + "|| ");
		                break;
		            case STRING:
		                System.out.print("Last evaluated as \"" + cell.getStringCellValue() + "|| ");
		                break;
		            case BOOLEAN:
		                System.out.print("Last evaluated as \"" + cell.getBooleanCellValue() + "|| ");
		                break;
		            case ERROR:
		            	System.out.print("Last evaluated as \"" + cell.getErrorCellValue() + "|| ");
		            default: System.out.print("FORMULA CELL NOT EVELUATED YET|| ");
		        	}
		        case BLANK: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print("|| ");
		            break;
		        case _NONE: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print("|| ");
		            break;
		        }
			}
			System.out.println();
		}
		
		System.out.println();
		System.out.println();
		System.out.println();
		
		//Find number of rows in excel file
	    int rowCount = myExcelBookSheet.getLastRowNum()-myExcelBookSheet.getFirstRowNum();

	    //Create a loop over all the rows of excel file to read it
	    for (int i = 0; i < rowCount+1; i++) {

	        Row rowForLoop = myExcelBookSheet.getRow(i);

	        //Create a loop to print cell values in a row
	        for (int j = 0; j < rowForLoop.getLastCellNum(); j++) {
	        	switch (rowForLoop.getCell(j,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellTypeEnum()) {
	            case STRING: 
	            	//Print Excel STRING data in console
		            System.out.print(rowForLoop.getCell(j).getStringCellValue()+"|| "); 
		            break;
	            case NUMERIC: 
	            	//Print Excel data in console
	            	if (DateUtil.isCellDateFormatted(rowForLoop.getCell(j))) {
	            		//Print Excel DATE data in dateFormat to console
			            System.out.print(rowForLoop.getCell(j).getDateCellValue() + "|| "); 
                    } else {
                    	//Print Excel NUMERIC data to console
			            System.out.print(rowForLoop.getCell(j).getNumericCellValue() + "|| ");
                    }
	            	break;
	            case BOOLEAN: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print(rowForLoop.getCell(j).getBooleanCellValue() + "|| ");
		            break;
		        case FORMULA: 
		        	//Print Excel BOOLEAN data to console
		        	switch(rowForLoop.getCell(j).getCachedFormulaResultTypeEnum()) {
		            case NUMERIC:
		            	System.out.print("Last evaluated as: " + rowForLoop.getCell(j).getNumericCellValue() + "|| ");
		                break;
		            case STRING:
		                System.out.print("Last evaluated as \"" + rowForLoop.getCell(j).getStringCellValue() + "|| ");
		                break;
		            case BOOLEAN:
		                System.out.print("Last evaluated as \"" + rowForLoop.getCell(j).getBooleanCellValue() + "|| ");
		                break;
		            case ERROR:
		            	System.out.print("Last evaluated as \"" + rowForLoop.getCell(j).getErrorCellValue() + "|| ");
		            default: System.out.print("FORMULA CELL NOT EVELUATED YET|| ");
		        	}
		        case BLANK: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print("|| ");
		            break;
		        case _NONE: 
	            	//Print Excel BOOLEAN data to console
		            System.out.print("|| ");
		            break;
		        }

	        }
	        System.out.println();
	    } 
		
//		collect in hashmap and pass the data whereever you want to file db etc
		/*Map<Integer, List<String>> data = new HashMap<>();
		int i = 0;
		for (Row row : myExcelBookSheet) {
		    data.put(i, new ArrayList<String>());
		    for (Cell cell : row) {
		        switch (cell.getCellType()) {
		            case STRING: ... break;
		            case NUMERIC: ... break;
		            case BOOLEAN: ... break;
		            case FORMULA: ... break;
		            default: data.get(new Integer(i)).add(" ");
		        }
		    }
		    i++;
		}*/
		
		
		inputStream.close();
		myExcelBook.close();
	}
}
