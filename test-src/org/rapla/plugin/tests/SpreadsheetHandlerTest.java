package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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
		
		SpreadsheetHandler sh = new SpreadsheetHandler();
		
		sh.createScheduleSheet("AMOD1", "12/10", "12/16");
		sh.createDateRow("AMOD1", 3, "Monday", "12/3/2018");
		sh.addShotToSchedule("AMOD1", 4, "Load & Cycle", "0600", "0900", "CDLMS1", "Kevin");
		
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
		assertTrue(sh.workbook.getSheet("AMOD1").getColumnWidth(10) == 6000);
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
		
		SpreadsheetHandler sh = new SpreadsheetHandler();
		
		sh.populateBulkUpload("Shot_Template", 
				  "BL10_SUITE", 
				  1, 
				  "TEST SHOT 1", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30B, LIVE CEC, TE: Element, BL10_SUITE, ELEMENT: OASIS, CONFIG: BL10_DDG", 
				  "Connelly, John");
		
		sh.populateBulkUpload("Shot_Template", 
				  "LBTS", 
				  2, 
				  "TEST SHOT 2", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, CDLMS2, TE: Element, LBTS, ELEMENT: CND, CONFIG: BL9_DDG", 
				  "Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "SUITE_B", 
				  3, 
				  "TEST SHOT 3", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, LIVE CEC, UMG1, TE: Integration, SUITE_B, ELEMENT: MA, CONFIG: BMD51_DDG", 
				  "Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "AMOD1", 
				  4, 
				  "TEST SHOT 4", 
				  "12/19/2018", 
				  "0600", 
				  "0900", 
				  "BUILD: 30, TE: SSIT, AMOD1, ELEMENT: WCS, CONFIG: BMD51_DDG", 
				  "Tooley, Kevin");
		
		sh.populateBulkUpload("Shot_Template", 
				  "TI16", 
				  5, 
				  "TEST SHOT 5", 
				  "12/19/2018", 
				  "1200", 
				  "1500", 
				  "BUILD: 30, TE: SSIT, TI16, ELEMENT: WCS, CONFIG: BL9_DDG", 
				  "Tooley, Kevin");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/BU_file1.xlsx", 
				"C:/Users/ktooley/Documents/TEST/BU_file2.xls");
		
		assertTrue(f.exists());
		assertTrue(f1.exists());
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(0).getStringCellValue().equals("Connelly, John"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(1).getStringCellValue().equals("TEST SHOT 1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(3).getStringCellValue().equals("USN-CSEA ACB20"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(5).getStringCellValue().equals("OASIS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(1).getCell(6).getStringCellValue().equals("12/19/2018"));
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
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(0).getStringCellValue().equals("Tooley, Kevin"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(1).getStringCellValue().equals("TEST SHOT 2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(4).getStringCellValue().equals("Element"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(5).getStringCellValue().equals("CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(6).getStringCellValue().equals("12/19/2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(9).getStringCellValue().equals("LBTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(10) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(11) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(12) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(13) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(14) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(15).getStringCellValue().equals("CDLMS2"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(16).getStringCellValue().equals("MLST3 (CDLMS2)"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(2).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(0).getStringCellValue().equals("Tooley, Kevin"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(1).getStringCellValue().equals("TEST SHOT 3"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(4).getStringCellValue().equals("Integration"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(5).getStringCellValue().equals("MA"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(6).getStringCellValue().equals("12/19/2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(9).getStringCellValue().equals("SUITE B"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(10) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(11) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(12) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(13) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(14) == null);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(15).getStringCellValue().equals("UMG1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(16).getStringCellValue().equals("UMG-1 SUPPORT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(3).getCell(17).getStringCellValue().equals("LIVE CEC/WASP"));

		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(0).getStringCellValue().equals("Tooley, Kevin"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(1).getStringCellValue().equals("TEST SHOT 4"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(3).getStringCellValue().equals("BMD5.1"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(6).getStringCellValue().equals("12/19/2018"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(7).getNumericCellValue() == 600);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(8).getNumericCellValue() == 900);
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(9).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 CND"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(10).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(11).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 SPY"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(12).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ADS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(13).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ACTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(14).getStringCellValue().equals("AMOD NSCC TI12 SUITE 1 ORTS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(15).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(16).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(4).getCell(17).getStringCellValue().equals(""));
		
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(0).getStringCellValue().equals("Tooley, Kevin"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(1).getStringCellValue().equals("TEST SHOT 5"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(2).getStringCellValue().equals(""));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(3).getStringCellValue().equals("USN-CSEA ACB16"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(4).getStringCellValue().equals("SSIT"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(5).getStringCellValue().equals("WCS"));
		assertTrue(sh.bulkUpload.getSheet("Shot_Template").getRow(5).getCell(6).getStringCellValue().equals("12/19/2018"));
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

}
