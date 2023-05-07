package exceloperations;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;										// For Rows & Cells of Iterator<?>
import org.apache.poi.xssf.usermodel.*;

public class ReadData {

	public static void main(String[] args) throws IOException {
		
//		Path path = FileSystems.getDefault().getPath("").toAbsolutePath();	// Get the working directory path
//		System.out.println(path);
		
		String excelFilePath = "./Data/NST_Batch_Status.xlsx";					// Path of excelSheet
		
		FileInputStream inputStream = new FileInputStream(excelFilePath);	// Creating input stream to access the excelSheet
		
		XSSFWorkbook workBook = new XSSFWorkbook(inputStream);				// Get the workBook
		
//		XSSFSheet sheet = workBook.getSheet("Sheet1");						// Access the sheet by Name
		
		XSSFSheet sheet = workBook.getSheetAt(0); 							// Access the sheet by Index
  
	/** Read data using Loop through count of Rows & Cells  **/		
/*		int rows = sheet.getLastRowNum();									// Get the last row number (Total Rows)
		int cols = sheet.getRow(1).getLastCellNum();						// Get the last column number (Total columns in a row)
		
//		System.out.println(rows +" " +cols);
		
		for(int r=0; r<=rows; r++) {										// Loop through the rows from beginning to End
			
			XSSFRow row = sheet.getRow(r);									// Object of each Row
			
			for(int c=0; c<cols; c++) {										// Loop through each cell in a Row
				
				XSSFCell cell = row.getCell(c);								// Object of each cell in a row
				
				switch(cell.getCellType()) {								// Getting type of the Cell data
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				default: System.out.println("Invalid Type");
				}
				
				System.out.print(" | ");
			}
			System.out.println();
		}
*/		
	/** Read data using Iterator Method **/
		Iterator<Row> rowIterator = sheet.iterator();						// Create a iterator to iterate Rows
		
		while(rowIterator.hasNext()) {										// Checking for next row available or not?
			XSSFRow row = (XSSFRow) rowIterator.next();						// Getting the object of row
			
			Iterator<Cell> cellIterator = row.iterator();					// Create a iterator to iterate Cells
			
			while(cellIterator.hasNext()) {									// Checking for next cells available or not?
				XSSFCell cell = (XSSFCell) cellIterator.next();				// Getting the object of row
				Date date = new Date();
				switch(cell.getCellType()) {								// Getting type of the Cell data
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print((int)cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				default: System.out.println("Invalid Type");
				}
				
				System.out.print(" | ");
			}
			System.out.println();
		}
		
		workBook.close();
	}

}
