package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Test;
import org.rapla.plugin.schedulegenie.SpreadsheetHandler;

import com.extentech.ExtenXLS.Cell;

import jxl.CellType;

public class SpreadsheetHandlerTest {

	@Test
	public void testCreateXlsxSpreadsheet() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/poi-generated-file.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		sh.createScheduleSheet("AMOD1", "12/10", "12/16");
		sh.createDateRow("AMOD1", 3, "Monday", "12/3/2018");
		sh.addShotToSchedule("AMOD1", 4, "Load & Cycle", "0600", "0900", "CDLMS1", "Tooley, Kevin");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/poi-generated-file.xlsx", "C:/Users/ktooley/Documents/TEST/poi-generated-file1.xlsx");
		
		assertTrue(f.exists());
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(0) == 2600);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(1) == 2150);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(2) == 11300);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(3) == 2400);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(4) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(5) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(6) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(7) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(8) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(9) == 1700);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(10) == 1700);
		//assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(10) == 6000);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(12) == 4800);
		
		assertTrue(sh.workbook.getSheet("AMOD1").getNumMergedRegions() == 2);
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(0).getCell(0).getCellStyle().getFont().getFontHeightInPoints() == 10);
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(0).getCellStyle().getFont().getFontHeightInPoints() == 8);
		
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(0).getStringCellValue().equals("Start Time"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(1).getStringCellValue().equals("End Time"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(2).getStringCellValue().equals("Element"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(3).getStringCellValue().equals("B/L"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(4).getStringCellValue().equals("CDLMS1"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(5).getStringCellValue().equals("CDLMS2"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(6).getStringCellValue().equals("UMG1"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(7).getStringCellValue().equals("UMG2"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(8).getStringCellValue().equals("CEC"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(9).getStringCellValue().equals("JMCIS"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(10).getStringCellValue().equals("MMSP"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(11).getStringCellValue().equals("Responsible Individual(s)"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(12).getStringCellValue().equals("Support"));
		
		
		// TODO:  Add test to check names of sheets
		
	}
	
	/**
	 * This test will populate the bulk upload spreadsheet with a standard eight shots per day
	 * beginning at 0000 on Monday morning and going through Sunday night at 2400.  
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	@Test
	public void testNEWPopulateBulkUpload() throws FileNotFoundException, IOException {
		
		/*
		 * This method does not actually use the BU_EveryShot1.xlsx spreadsheet.  However,
		 * the closeWorkbook() method does require that both the xlsx and xls spreadsheets
		 * as arguments.  
		 */
		File f = new File("C:/Users/ktooley/Documents/TEST/BU_EveryShot1.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		File f1 = new File("C:/Users/ktooley/Documents/TEST/BU_EveryShot2.xls");
		if (f1.exists()) {
			if (f1.delete()) { System.out.println("File deleted"); }
		}
		
		// The following are input arguments for the populateBulkUpload method
		int dayInt = 1;   // day number in date argument 
		int shotInt = 1;  // shot number in shotName argument 
		int rowInt = 1;   // row number in rowNumber argument 
		
		// This keeps track of array index for the shotStartEndTimes array
		int startCounter = 0; 
		String [] shotStartEndTimes = {"0000", "0300", "0600", "0900", "1200", "1500", "1800", "2100", "2400"};
		
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		// For loop for each day of week
		for (int i = 0; i < 7; i++) {
						
			// For loop for each shot of day
			for (int j = 0; j < 8; j++) {
				
				sh.populateBulkUpload("Shot_Template", 
						  "BL10_SUITE", 
						  rowInt, 
						  "TEST SHOT " + shotInt, 
						  "1/" + dayInt + "/2019", 
						  shotStartEndTimes[startCounter], 
						  shotStartEndTimes[startCounter + 1], 
						  "BUILD: 30B, CDLMS2, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG, LIVE MMSP", 
						  "Tooley, Kevin");
				
				shotInt++;
				rowInt++;
				startCounter++;
				
			}
			
			dayInt++;
			startCounter = 0;  // Resets at the end of the day
			
		}
		
		/*
		 * See note above regarding the closeWorkbook() method.  
		 * Summary:  Although we don't initialize both files, we need both in this call.
		 */
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/BU_EveryShot1.xlsx", 
				"C:/Users/ktooley/Documents/TEST/BU_EveryShot2.xls");
		
		assertTrue(f.exists());
		assertTrue(f1.exists());
		
		// Reset all counters in anticipation for asserts
		dayInt = 1;
		shotInt = 1;
		rowInt = 1;
		startCounter = 0;
		
		// For loop for each day of week
		for (int i = 0; i < 7; i++) {
						
			// For loop for each shot of day
			for (int j = 0; j < 8; j++) {
				
				String tmpDay = String.format("%02d", dayInt);  // Forces a two digit day number for the toString() of getCell(6) below
				
				/*
				 * Asserts for each cell in row.
				 * 
				 * Note:  At row 36 of the spreadsheet (36 as labeled in Excel), the date field turns into an integer representing the
				 * date.  This appears to function without issue on the Test Site Scheduling System, but it is visually unappealing.
				 * Troubleshooting the local application does not show any errors.  I suspect that Microsoft may have an issue with
				 * this. 
				 */
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(0).getStringCellValue().equals("Tooley, Kevin"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(1).getStringCellValue().equals("TEST SHOT " + shotInt));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(2).getStringCellValue().equals(""));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(3).getStringCellValue().equals("USN-CSEA ACB20"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(4).getStringCellValue().equals("Element"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(5).getStringCellValue().equals("SYS ADMIN"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(6).toString().equals(tmpDay + "-Jan-2019")); //.format("%02d\n", i);  // Will print 09
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(7).getNumericCellValue() == Integer.parseInt(shotStartEndTimes[startCounter]));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(8).getNumericCellValue() == Integer.parseInt(shotStartEndTimes[startCounter + 1]));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(9).getStringCellValue().equals("NSCC BL10 CND"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(10).getStringCellValue().equals("NSCC BL10 WCS"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(11).getStringCellValue().equals("NSCC BL10 SPY"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(12).getStringCellValue().equals("NSCC BL10 ADS"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(13).getStringCellValue().equals("NSCC BL10 ACTS"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(14).getStringCellValue().equals("NSCC BL10 ORTS"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(15).getStringCellValue().equals("CDLMS2"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(16).getStringCellValue().equals("MLST3 (CDLMS2)"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));
				assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(rowInt).getCell(18).getStringCellValue().equals("LIVE MMSP"));
			
				shotInt++;
				rowInt++;
				startCounter++;
				
			}
			
			dayInt++;
			startCounter = 0;  // Resets at the end of the day
			
		}

	}
	
	/**
	 * This tests all the different quantities of Responsible Individual (RI) names.  
	 * 
	 * Switch statement based on the array length.  
     * Based on the constraints that there will alway be a first and last name for each person (Rapla Requirement), the
     * following algorithms format the string in a "First Last, Suffix" syntax with a " | " separator. 
     * 
     * Here is an example of the logic table used to evaluate the algorithm logic for an arrayLength of 7:
     * 
     * "L" = last name
     * "F" = first name
     * "x" = a suffix like Jr.
     * 
     * 0 | 1 | 2 | 3 | 4 | 5 | 6 
     * L   x   F   L   F   L   F
     * L   F   L   x   F   L   F
     * L   F   L   F   L   x   F
     * 
     * No other possibilities exist with given constraints.
     * 
     * Array length of 9 through 12 will warn the operator to use less shot RI's
     * 
     * Any other array lengths warn the operator to check the RI fields.
     * 
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	@Test
	public void testGetShotRiString() throws FileNotFoundException, IOException {
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		//Size 2
		String ri1 = "Tooley, Kevin";
		String test1 = sh.getShotRiString(ri1);
		assertTrue(test1.equals("Kevin Tooley"));
		
		//Size 3
		String ri2 = "Tooley, Jr., Kevin";
		String test2 = sh.getShotRiString(ri2);
		assertTrue(test2.equals("Kevin Tooley, Jr."));
		
		//Size 4
		String ri3 = "Tooley, Kevin, Connelly, John";
		String test3 = sh.getShotRiString(ri3);
		assertTrue(test3.equals("Kevin Tooley | John Connelly"));
		
		//Size 5 (Last, Jr., First | Last, First) or (Last, First | Last, Jr., First)
		String ri4 = "Tooley, Jr., Kevin, Connelly, John";
		String test4 = sh.getShotRiString(ri4);
		assertTrue(test4.equals("Kevin Tooley, Jr. | John Connelly"));
		
		String ri5 = "Tooley, Kevin, Connelly, Jr., John";
		String test5 = sh.getShotRiString(ri5);
		assertTrue(test5.equals("Kevin Tooley | John Connelly, Jr."));
		
		//Size 6 (i.e. two Jr's, or 3 regular RI's)
		String ri6 = "Tooley, Kevin, Connelly, John, Donow, Matthew";
		String test6 = sh.getShotRiString(ri6);
		assertTrue(test6.equals("Kevin Tooley | John Connelly | Matthew Donow"));
		
		String ri7 = "Tooley, Sr., Kevin, Connelly, Jr., John";
		String test7 = sh.getShotRiString(ri7);
		assertTrue(test7.equals("Kevin Tooley, Sr. | John Connelly, Jr."));
		
		//Size 7 (i.e. one Jr, or 2 regular RI's)
		String ri8 = "Tooley, Jr., Kevin, Connelly, John, Donow, Matthew";
		String test8 = sh.getShotRiString(ri8);
		assertTrue(test8.equals("Kevin Tooley, Jr. | John Connelly | Matthew Donow"));
		
		String ri9 = "Tooley, Kevin, Connelly, Sr., John, Donow, Matthew";
		String test9 = sh.getShotRiString(ri9);
		assertTrue(test9.equals("Kevin Tooley | John Connelly, Sr. | Matthew Donow"));
		
		String ri10 = "Tooley, Kevin, Connelly, John, Donow, Jr., Matthew";
		String test10 = sh.getShotRiString(ri10);
		assertTrue(test10.equals("Kevin Tooley | John Connelly | Matthew Donow, Jr."));
		
		//Size 8 (i.e. two Jr, or 1 regular RI, or no Jr's at all)
		String ri11 = "Tooley, Kevin, Connelly, John, Donow, Matthew, Frank, Dave";
		String test11 = sh.getShotRiString(ri11);
		assertTrue(test11.equals("Kevin Tooley | John Connelly | Matthew Donow | Dave Frank"));
		
		String ri12 = "Tooley, Jr., Kevin, Connelly, Sr., John, Donow, Matthew";
		String test12 = sh.getShotRiString(ri12);
		assertTrue(test12.equals("Kevin Tooley, Jr. | John Connelly, Sr. | Matthew Donow"));
		
		String ri13 = "Tooley, Jr., Kevin, Connelly, John, Donow, Sr., Matthew";
		String test13 = sh.getShotRiString(ri13);
		assertTrue(test13.equals("Kevin Tooley, Jr. | John Connelly | Matthew Donow, Sr."));
		
		String ri14 = "Tooley, Kevin, Connelly, Jr., John, Donow, Sr., Matthew";
		String test14 = sh.getShotRiString(ri14);
		assertTrue(test14.equals("Kevin Tooley | John Connelly, Jr. | Matthew Donow, Sr."));
		
		//ERRORS
		String ri15 = "Tooley, Kevin, Connelly, John, Donow, Matthew, Frank, Dave, Tooley, Kevin, Connelly, John";
		String test15 = sh.getShotRiString(ri15);
		assertTrue(test15.equals("NAME ERROR: EXCEEDED THE NUMBER OF SHOT OWNERS"));
		
		String ri16 = "Tooley";
		String test16 = sh.getShotRiString(ri16);
		assertTrue(test16.equals("NAME ERROR: CHECK INPUTS"));
		
	}
	
	@Test
	public void testXSSFDateFieldOnSpreadsheet() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/test.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		// Create a Sheet
        XSSFSheet sheet = sh.workbook.createSheet("TestDate");
        
        // Create a Font for styling new row
        XSSFFont newRowFont = sh.workbook.createFont();
        newRowFont.setFontName("ARIAL");
        newRowFont.setFontHeightInPoints((short) 9);
        newRowFont.setBold(false);
        
        // Create a Font for styling time fields in new row
        XSSFFont newRowTimeFont = sh.workbook.createFont();
        newRowTimeFont.setFontName("ARIAL");
        newRowTimeFont.setFontHeightInPoints((short) 8);
        newRowTimeFont.setBold(false);
        
        // Create a CellStyle with the font
        XSSFCellStyle newRowCellStyle = sh.workbook.createCellStyle();
        newRowCellStyle.setFont(newRowFont);
        newRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newRowCellStyle.setBorderBottom(BorderStyle.THIN);
        newRowCellStyle.setBorderTop(BorderStyle.THIN);
        newRowCellStyle.setBorderRight(BorderStyle.THIN);
        newRowCellStyle.setBorderLeft(BorderStyle.THIN);
        
        // Create a CellStyle with the font for the time fields
        DataFormat format = sh.workbook.createDataFormat(); // Sets up format for the time fields
        XSSFCellStyle newRowTimeCellStyle = sh.workbook.createCellStyle();
        newRowTimeCellStyle.setFont(newRowTimeFont);
        newRowTimeCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newRowTimeCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newRowTimeCellStyle.setBorderBottom(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderTop(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderRight(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderLeft(BorderStyle.THIN);
        newRowTimeCellStyle.setDataFormat(format.getFormat("0000"));
        
        String cellValue = "12/19/2018";
        
        for (int i = 0; i < 40; i++) {
        	// Add row to sheet
            XSSFRow newRow = sheet.createRow(i);
            
            XSSFCell cell = newRow.createCell(0);
            
            java.util.Date datetemp = null;
        	SimpleDateFormat format1 = new SimpleDateFormat("M/d/yy");
        	try {
				datetemp = format1.parse(cellValue);
			} catch (ParseException e) {
				// Auto-generated catch block
				e.printStackTrace();
			}
        	
        	//cell.setCellType(CellType.NUMERIC);
        	//cell.setCellType(CellType.DATE);
        	XSSFCellStyle style = sh.workbook.createCellStyle();
        	style.setDataFormat(14);
        	cell.setCellValue(datetemp);
        	cell.setCellStyle(style);
        }
        
        FileOutputStream testOutStream = null;
        
        try {
			if (sh.isUnitTest())
				testOutStream = new FileOutputStream("C:/Users/ktooley/Documents/TEST/test.xlsx");
			//else
			//	testOutStream = new FileOutputStream(chooseFile("C:/Users/ktooley/Documents/TEST/test.xlsx"));
			//bulkOutStream = new FileOutputStream(bulkFilePath);
		} catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			sh.workbook.write(testOutStream);
			//bulkUpload.write(bulkOutStream);
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			testOutStream.close();
			//bulkOutStream.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			sh.workbook.close();
			//bulkUpload.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	@Test
	public void testHSSFDateFieldOnSpreadsheet() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/test.xls");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		// Create a Sheet
        HSSFSheet sheet = workbook.createSheet("TestDate");
        
        String cellValue = "12/19/2018";
        
        for (int i = 0; i < 40; i++) {
        	HSSFRow newRow = sheet.createRow(i);
        	
        	HSSFCell cell = newRow.createCell(0);
        	
        	java.util.Date datetemp = null;
        	SimpleDateFormat format = new SimpleDateFormat("M/d/yy");
        	try {
				datetemp = format.parse(cellValue);
			} catch (ParseException e) {
				// Auto-generated catch block
				e.printStackTrace();
			}
        	
        	//cell.setCellType(CellType.NUMERIC);
        	HSSFCellStyle style = workbook.createCellStyle();
        	style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy"));
        	cell.setCellValue(datetemp);
        	cell.setCellStyle(style);
        }
        
        FileOutputStream testOutStream = null;
        
        try {
			//if (sh.isUnitTest())
				testOutStream = new FileOutputStream("C:/Users/ktooley/Documents/TEST/test.xls");
			//else
			//	testOutStream = new FileOutputStream(chooseFile("C:/Users/ktooley/Documents/TEST/test.xlsx"));
			//bulkOutStream = new FileOutputStream(bulkFilePath);
		} catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			workbook.write(testOutStream);
			//bulkUpload.write(bulkOutStream);
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			testOutStream.close();
			//bulkOutStream.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			workbook.close();
			//bulkUpload.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	@Test
	public void TestUpdateYesNoBox() {
		
		try {
			
			// Initialize the handler
			SpreadsheetHandler sh = new SpreadsheetHandler(true);
			
			// Test 1:  Init result bool and set with dialog box
			boolean result = sh.UpdateYesNoBox();
			
			// Print the operator choice
			if (result)
				System.out.println("User pressed YES");
			else
				System.out.println("User pressed NO");
			
			Thread.sleep(1000);
			
			// Test 2:  set result bool again
			result = sh.UpdateYesNoBox();
			
			if (result)
				System.out.println("User pressed YES");
			else
				System.out.println("User pressed NO");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

	}

}
