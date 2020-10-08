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
import java.util.ArrayList;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.rapla.plugin.metricsgenie.MetricsSpreadsheetHandler;

import com.extentech.ExtenXLS.Cell;

import jxl.CellType;

public class MetricsSpreadsheetHandlerTest {

	@Test
	public void testCreateXlsxSpreadsheet() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/metric-poi-generated-file.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		MetricsSpreadsheetHandler sh = new MetricsSpreadsheetHandler(true);
		
		sh.createScheduleSheet();
		//sh.createDateRow("AMOD1", 3, "Monday", "12/3/2018");
		//sh.addShotToSchedule("AMOD1", 4, "Load & Cycle", "0600", "0900", "CDLMS1", "Tooley, Kevin");
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/metric-poi-generated-file.xlsx");
		
		assertTrue(f.exists());
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(0) == 3900);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(1) == 4100);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(2) == 7500);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(3) == 3800);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(4) == 5700);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(5) == 4200);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(6) == 3800);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(7) == 3800);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(8) == 5100);
		assertTrue(sh.workbook.getSheet("Schedule").getColumnWidth(9) == 4200);
		
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(0).getStringCellValue().equals("ELEMENT"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(1).getStringCellValue().equals("PROGRAM"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(2).getStringCellValue().equals("FUNDING SOURCE"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(3).getStringCellValue().equals("BUILD"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(4).getStringCellValue().equals("EFFORT"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(5).getStringCellValue().equals("SYSTEM"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(6).getStringCellValue().equals("START DATE"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(7).getStringCellValue().equals("END DATE"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(8).getStringCellValue().equals("TOTAL DURATION"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(0).getCell(9).getStringCellValue().equals("USER"));
		
		
		// TODO:  Add test to check names of sheets
		
	}
	
	@Test
	public void testAddShotToSchedule() throws FileNotFoundException, IOException {
		
		File f = new File("C:/Users/ktooley/Documents/TEST/metric-AddShotToSchedule_Test.xlsx");
		if (f.exists()) {
			if (f.delete()) { System.out.println("File deleted"); }
		}
		
		MetricsSpreadsheetHandler sh = new MetricsSpreadsheetHandler(true);
		
		sh.createScheduleSheet();

		ArrayList<Object> shotData = new ArrayList<Object>();
		shotData.add("CND:CPM:TI16");
		shotData.add("10/6/20 5:00 AM");
		shotData.add("10/6/20 8:30 AM");
		shotData.add("*SWCR, MES-8, F:CPM-FY19, C:9C22D, Build 32B, +CND");
		shotData.add("Schraut, Karl");
		shotData.add("3,30");
		
		sh.addShotToSchedule("DT4", 1, shotData);
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/metric-AddShotToSchedule_Test.xlsx");
		
		assertTrue(f.exists());
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(0).getStringCellValue().equals("CND"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(1).getStringCellValue().equals("9C22D"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(2).getStringCellValue().equals("CPM-FY19"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(3).getStringCellValue().equals("32B"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(4).getStringCellValue().equals("SWCR"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(5).getStringCellValue().equals("DT4"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(6).getStringCellValue().equals("10/6/20 5:00 AM"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(7).getStringCellValue().equals("10/6/20 8:30 AM"));
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(8).getNumericCellValue() == 3.5);
		assertTrue(sh.workbook.getSheet("Schedule").getRow(1).getCell(9).getStringCellValue().equals("Schraut"));
		
		
		
		
		// TODO:  Add test to check names of sheets
		
	}

}
