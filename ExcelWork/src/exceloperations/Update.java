package exceloperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.excelant.ExcelAntSetDoubleCell;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xml.security.configuration.SecurityHeaderHandlersType;


public class Update {

public static void main(String[] args) throws IOException {
		
//		String excelFilePath = "./Data/NST_Batch_Status1.xlsx";				// Path of excelSheet
		
		String excelFilePath = "https://1drv.ms/x/s!Aj42DlSY2lUInTu0-Uuj0w-B4g8Z?e=alHH52";
		
		FileInputStream inputStream = new FileInputStream(excelFilePath);	// Creating input stream to access the excelSheet
		
		XSSFWorkbook workBook = new XSSFWorkbook(inputStream);				// Get the workBook
		
		XSSFSheet sheet = workBook.getSheetAt(0);							// Get a Sheet at Index 0
		
		/** Using For Loop **/
		int rows = sheet.getLastRowNum();									// Length of Rows
		int cells = sheet.getRow(rows).getLastCellNum();					// Length of Cells in each row
		
		System.out.println(rows + " " + cells);
		
		XSSFRow row = sheet.getRow(rows);
		
//		Date date = new Date();
//		String month = date.toString();		//Sun May 07 17:23:19 IST 2023
		
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat month = new SimpleDateFormat("MMM");
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
//	    System.out.println("Date and time = "+dateFormat.format(cal.getTime()));
	    
		String startTime = "16:00:00";
		String endTime = "16:34:12";
		int status = 1;
//		String mString  = dateFormat.format(cal.getTime());
		
	    
		XSSFCell cell;
		
		cell = row.getCell(1);
//		String cString = cell.getStringCellValue();
//		System.out.println(mString + "\n" + cString);
//		System.out.println(mString.equals(cString));
		if((cell.getStringCellValue()).equals(month.format(cal.getTime()))) {
			cell = row.getCell(2);
			if((cell.getStringCellValue()).equals(dateFormat.format(cal.getTime()))) {
//				int setCell=0;
//				for(int c=3; c<cells; c++) {
//					cell = row.getCell(c);
//					if(cell.getCellType() == null) {
//						setCell = c;
//						break;
//					}
//				}
			
				cell = row.createCell(cells++);						// Start Time
				cell.setCellValue(startTime);
				
				cell = row.createCell(cells++);						// End Time
				cell.setCellValue(endTime);
				
				cell = row.createCell(cells++);						// Status
				cell.setCellValue((status==1? "Success":"Failed"));
				
				cell = row.createCell(cells++);						// Time Duration
				cell.setCellValue(endTime.substring(0, 5));
			}	
			else {	
				rows++;
				cells = 0;
				row = sheet.createRow(rows);						// New day create new Row
				
				cell = row.createCell(cells++);						// Sr No.
				cell.setCellValue(sheet.getRow(rows-1).getCell(0).getNumericCellValue()+1);
				
				cell = row.createCell(cells++);						// Month
				cell.setCellValue(sheet.getRow(rows-1).getCell(1).getStringCellValue());
				
				cell = row.createCell(cells++);						// Date
				cell.setCellValue(dateFormat.format(cal.getTime()));
				
				cell = row.createCell(cells++);						// Start Time
				cell.setCellValue(startTime);
				
				cell = row.createCell(cells++);						// End Time
				cell.setCellValue(endTime);
				
				cell = row.createCell(cells++);						// Status
				cell.setCellValue((status==1? "Success":"Failed"));
				
				cell = row.createCell(cells++);						// Time Duration
				cell.setCellValue(endTime.substring(0, 5)); 
					
				
			}
		}
		else {
			rows++;
			cells = 0;
			row = sheet.createRow(rows);						// New day create new Row
			
			cell = row.createCell(cells++);						// Sr No.
			cell.setCellValue(sheet.getRow(rows-1).getCell(0).getNumericCellValue()+1);
			
			cell = row.createCell(cells++);						// Month
			cell.setCellValue(sheet.getRow(rows-1).getCell(1).getStringCellValue());
			
			cell = row.createCell(cells++);						// Date
			cell.setCellValue(dateFormat.format(cal.getTime()));
			
			cell = row.createCell(cells++);						// Start Time
			cell.setCellValue(startTime);
			
			cell = row.createCell(cells++);						// End Time
			cell.setCellValue(endTime);
			
			cell = row.createCell(cells++);						// Status
			cell.setCellValue((status==1? "Success":"Failed"));
			
			cell = row.createCell(cells++);						// Time Duration
			cell.setCellValue(endTime.substring(0, 5));
		}
	
		inputStream.close();												// Closing input stream
		
		FileOutputStream outputStream = new FileOutputStream(excelFilePath);// Creating object for output Stream of excel file
		
		workBook.write(outputStream);										// Writing data into Excel sheet (Getting the data from all the objects which is created)
		workBook.close();													// Closing workbook object
		outputStream.close();												// Closing output stream
		
		System.out.println("Successful...");
	}

}
