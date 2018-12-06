package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;

import org.junit.Test;
import org.rapla.plugin.schedulegenie.SpreadsheetHandler;

public class SpreadsheetHandlerTest {

	@Test
	public void testCreateXlsxSpreadsheet() {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/poi-generated-file.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		SpreadsheetHandler sh = new SpreadsheetHandler();
		
		sh.createScheduleWorkbook();
		sh.createDateRow("AMOD1", 3, "Monday", "12/3/2018");
		sh.addShotToSchedule("AMOD1", 4, 600, 900, "Load & Cycle");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/poi-generated-file.xlsx");
		
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

}
