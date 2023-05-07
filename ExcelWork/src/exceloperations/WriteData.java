package exceloperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;
import org.bouncycastle.asn1.x509.qualified.TypeOfBiometricData;

public class WriteData {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath = "./Data/JavaExcel.xlsx";						// Path of excelSheet
		
		FileInputStream inputStream = new FileInputStream(excelFilePath);	// Creating input stream to access the excelSheet
		
		XSSFWorkbook workBook = new XSSFWorkbook(inputStream);				// Get the workBook
		
		XSSFSheet sheet = workBook.createSheet("Sheet2");					// Create a new Sheet
		
		Object empData[][] = { {"Emp Id", "Name", "Destination"},			// Object of data to write in Sheet
							   {101, "Lokesh Parihar", "System Engineer"},
							   {102, "Mahendra Bishnoi", "Software Developer 2"},
							   {103, "Neeraj Kumar", "Senior Software Developer"}
							};
		
		/** Using For Loop **/
		int rows = empData.length;											// Length of Rows
		int cells = empData[0].length;										// Length of Cells in each row
		
		for(int r=0; r<rows; r++) {											// Loop for Rows
			XSSFRow row = sheet.createRow(r);								// Create a new row
			
			for(int c=0; c<cells; c++) {									// Loop for Cells
				XSSFCell cell = row.createCell(c);							// Create a new cell for each row, it will create one cell at a time
				
				Object value = empData[r][c];								// Getting one value at time
				
				if (value instanceof String) {								// Getting type of object
					cell.setCellValue((String) value); 						// Setting data for that cell in Cell Object
				}
				else if (value instanceof Integer) {						
					cell.setCellValue((Integer) value);
				}
				else if (value instanceof Boolean) {
					cell.setCellValue((Boolean) value);
				}
			}
		}
		
		inputStream.close();												// Closing input stream
		
		FileOutputStream outputStream = new FileOutputStream(excelFilePath);// Creating object for output Stream of excel file
		
		workBook.write(outputStream);										// Writing data into Excel sheet (Getting the data from all the objects which is created)
		workBook.close();													// Closing workbook object
		outputStream.close();												// Closing output stream

	}

}
