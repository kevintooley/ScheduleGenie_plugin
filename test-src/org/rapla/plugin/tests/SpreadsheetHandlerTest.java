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
		//assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(10) == 6000);
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(11) == 4800);
		
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
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(10).getStringCellValue().equals("Responsible Individual(s)"));
		assertTrue(sh.workbook.getSheet("AMOD1").getRow(2).getCell(11).getStringCellValue().equals("Support"));
		
		
		// TODO:  Add test to check names of sheets
		
	}
	
	@Test
	public void testPopulateBulkUpload() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/BU_file1.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		File f1 = new File("C:/Users/ktooley/Documents/TEST/BU_file2.xls");
		if (f1.exists()) {
			if (f1.delete()) { System.out.println("File deleted"); }
		}
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  1, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  2, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  3, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0000", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  4, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "2100", 
				  "2400", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  5, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/BU_file1.xlsx", 
				"C:/Users/ktooley/Documents/TEST/BU_file2.xls");
		
		assertTrue(f.exists());
		assertTrue(f1.exists());
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(1).getStringCellValue().equals("TEST SHOT 1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(3).getStringCellValue().equals("USN-CSEA ACB20"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(5).getStringCellValue().equals("SYS ADMIN"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(9).getStringCellValue().equals("NSCC BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(10).getStringCellValue().equals("NSCC BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(11).getStringCellValue().equals("NSCC BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(12).getStringCellValue().equals("NSCC BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(13).getStringCellValue().equals("NSCC BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(14).getStringCellValue().equals("NSCC BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(0).getStringCellValue().equals("Romvary, Jr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(1).getStringCellValue().equals("TEST SHOT 2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(5).getStringCellValue().equals("CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(9).getStringCellValue().equals("LBTS BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(10).getStringCellValue().equals("LBTS BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(11).getStringCellValue().equals("LBTS BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(12).getStringCellValue().equals("LBTS BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(13).getStringCellValue().equals("LBTS BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(14).getStringCellValue().equals("LBTS BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(15).getStringCellValue().equals("CDLMS2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(16).getStringCellValue().equals("MLST3 (CDLMS2)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(1).getStringCellValue().equals("TEST SHOT 3"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(4).getStringCellValue().equals("Integration"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(5).getStringCellValue().equals("MA"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(7).getNumericCellValue() == 0);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(9).getStringCellValue().equals("SUITE B"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(10).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(11).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(12).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(13).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(14).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(15).getStringCellValue().equals("UMG1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(16).getStringCellValue().equals("UMG-1 SUPPORT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));

		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(0).getStringCellValue().equals("Romvary, Sr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(1).getStringCellValue().equals("TEST SHOT 4"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(7).getNumericCellValue() == 2100);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(8).getNumericCellValue() == 2400);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(9).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(10).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(11).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(12).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(13).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(14).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(0).getStringCellValue().equals("Boegly, Leo (Jerry)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(1).getStringCellValue().equals("TEST SHOT 5"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(7).getNumericCellValue() == 1200);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(8).getNumericCellValue() == 1500);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(9).getStringCellValue().equals("NSCC TI16 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(10).getStringCellValue().equals("NSCC TI16 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(11).getStringCellValue().equals("NSCC TI16 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(12).getStringCellValue().equals("NSCC TI16 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(13).getStringCellValue().equals("NSCC TI16 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(14).getStringCellValue().equals("NSCC TI16 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(17).getStringCellValue().equals(""));
		
		
		
	}
	
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
	public void testExtendedPopulateBulkUpload() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/BU_file3.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		File f1 = new File("C:/Users/ktooley/Documents/TEST/BU_file4.xls");
		if (f1.exists()) {
			if (f1.delete()) { System.out.println("File deleted"); }
		}
		
		SpreadsheetHandler sh = new SpreadsheetHandler(true);
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  1, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  2, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  3, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  4, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  5, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  6, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  7, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  8, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  9, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  10, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  11, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  12, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  13, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  14, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  15, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  16, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  17, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  18, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  19, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  20, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  21, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  22, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  23, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  24, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  25, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  26, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  27, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  28, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  29, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  30, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  31, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  32, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  33, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  34, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  35, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  36, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: SYS ADMIN, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  37, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Romvary, Jr., Christian, Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  38, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Connelly, John, Romvary, Jr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  39, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Romvary, Sr., Christian");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  40, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Boegly, Leo (Jerry), Tooley, Kevin");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/BU_file3.xlsx", 
				"C:/Users/ktooley/Documents/TEST/BU_file4.xls");
		
		assertTrue(f.exists());
		assertTrue(f1.exists());
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(1).getStringCellValue().equals("TEST SHOT 1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(3).getStringCellValue().equals("USN-CSEA ACB20"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(5).getStringCellValue().equals("SYS ADMIN"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(9).getStringCellValue().equals("NSCC BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(10).getStringCellValue().equals("NSCC BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(11).getStringCellValue().equals("NSCC BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(12).getStringCellValue().equals("NSCC BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(13).getStringCellValue().equals("NSCC BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(14).getStringCellValue().equals("NSCC BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(31).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(0).getStringCellValue().equals("Romvary, Jr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(1).getStringCellValue().equals("TEST SHOT 2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(5).getStringCellValue().equals("CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(9).getStringCellValue().equals("LBTS BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(10).getStringCellValue().equals("LBTS BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(11).getStringCellValue().equals("LBTS BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(12).getStringCellValue().equals("LBTS BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(13).getStringCellValue().equals("LBTS BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(14).getStringCellValue().equals("LBTS BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(15).getStringCellValue().equals("CDLMS2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(16).getStringCellValue().equals("MLST3 (CDLMS2)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(32).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(1).getStringCellValue().equals("TEST SHOT 3"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(4).getStringCellValue().equals("Integration"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(5).getStringCellValue().equals("MA"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(9).getStringCellValue().equals("SUITE B"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(10).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(11).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(12).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(13).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(14).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(15).getStringCellValue().equals("UMG1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(16).getStringCellValue().equals("UMG-1 SUPPORT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(33).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));

		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(0).getStringCellValue().equals("Romvary, Sr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(1).getStringCellValue().equals("TEST SHOT 4"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(9).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(10).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(11).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(12).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(13).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(14).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(34).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(0).getStringCellValue().equals("Boegly, Leo (Jerry)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(1).getStringCellValue().equals("TEST SHOT 5"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(7).getNumericCellValue() == 1200);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(8).getNumericCellValue() == 1500);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(9).getStringCellValue().equals("NSCC TI16 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(10).getStringCellValue().equals("NSCC TI16 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(11).getStringCellValue().equals("NSCC TI16 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(12).getStringCellValue().equals("NSCC TI16 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(13).getStringCellValue().equals("NSCC TI16 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(14).getStringCellValue().equals("NSCC TI16 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(35).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(1).getStringCellValue().equals("TEST SHOT 1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(3).getStringCellValue().equals("USN-CSEA ACB20"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(5).getStringCellValue().equals("SYS ADMIN"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(9).getStringCellValue().equals("NSCC BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(10).getStringCellValue().equals("NSCC BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(11).getStringCellValue().equals("NSCC BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(12).getStringCellValue().equals("NSCC BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(13).getStringCellValue().equals("NSCC BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(14).getStringCellValue().equals("NSCC BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(36).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(0).getStringCellValue().equals("Romvary, Jr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(1).getStringCellValue().equals("TEST SHOT 2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(5).getStringCellValue().equals("CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(9).getStringCellValue().equals("LBTS BL10 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(10).getStringCellValue().equals("LBTS BL10 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(11).getStringCellValue().equals("LBTS BL10 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(12).getStringCellValue().equals("LBTS BL10 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(13).getStringCellValue().equals("LBTS BL10 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(14).getStringCellValue().equals("LBTS BL10 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(15).getStringCellValue().equals("CDLMS2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(16).getStringCellValue().equals("MLST3 (CDLMS2)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(37).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(1).getStringCellValue().equals("TEST SHOT 3"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(4).getStringCellValue().equals("Integration"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(5).getStringCellValue().equals("MA"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(9).getStringCellValue().equals("SUITE B"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(10).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(11).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(12).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(13).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(14).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(15).getStringCellValue().equals("UMG1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(16).getStringCellValue().equals("UMG-1 SUPPORT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(38).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));

		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(0).getStringCellValue().equals("Romvary, Sr., Christian"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(1).getStringCellValue().equals("TEST SHOT 4"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(9).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(10).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(11).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(12).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(13).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(14).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(39).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(0).getStringCellValue().equals("Boegly, Leo (Jerry)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(1).getStringCellValue().equals("TEST SHOT 5"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(6).toString().equals("19-Dec-2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(6).getCellStyle().getDataFormatString().equals("m/d/yy"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(6).getCellStyle().getDataFormat() == 14);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(7).getNumericCellValue() == 1200);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(8).getNumericCellValue() == 1500);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(9).getStringCellValue().equals("NSCC TI16 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(10).getStringCellValue().equals("NSCC TI16 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(11).getStringCellValue().equals("NSCC TI16 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(12).getStringCellValue().equals("NSCC TI16 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(13).getStringCellValue().equals("NSCC TI16 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(14).getStringCellValue().equals("NSCC TI16 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(40).getCell(17).getStringCellValue().equals(""));
		
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

}
