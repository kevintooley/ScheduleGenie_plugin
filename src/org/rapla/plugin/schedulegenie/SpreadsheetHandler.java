package org.rapla.plugin.schedulegenie;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import jxl.*;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException; 

public class SpreadsheetHandler {
	
	// Declare the workbook used for the lab schedules
	public XSSFWorkbook workbook;
	public XSSFWorkbook bulkUpload;
	
	//Constructor
	public SpreadsheetHandler() throws FileNotFoundException, IOException {
		// Create a Workbook for Lab schedules
        workbook = new XSSFWorkbook(); // new XSSFWorkbook() for generating `.xlsx` file
        
        //final String userHome = System.getProperty("user.home");
        //String filePath = userHome + "\\Documents\\nscc_bulk_template.xls";
        
        // Create workbook for bulk upload
        //bulkUpload = new XSSFWorkbook(new FileInputStream(filePath));
	}
	
	public void createScheduleSheet(String sheetName, String weekStartDate, String weekEndDate) {

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        XSSFCreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        XSSFSheet sheet = workbook.createSheet(sheetName);
        
        /*
         * 
         * Create Header Row A
         * 
         */
        // Create a Font for styling header cells
        XSSFFont headerRowAFont = workbook.createFont();
        headerRowAFont.setBold(true);
        headerRowAFont.setFontName("ARIAL");
        headerRowAFont.setFontHeightInPoints((short) 10);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        // TODO: Make method for CellStyle setup
        // Create a CellStyle with the font
        XSSFCellStyle headerRowACellStyle = workbook.createCellStyle();
        headerRowACellStyle.setFont(headerRowAFont);
        headerRowACellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerRowACellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerRowACellStyle.setBorderBottom(BorderStyle.THIN);
        headerRowACellStyle.setBorderTop(BorderStyle.THIN);
        headerRowACellStyle.setBorderRight(BorderStyle.THIN);
        headerRowACellStyle.setBorderLeft(BorderStyle.THIN);

        // Create Row A, merge, adjust column widths
        XSSFRow headerRowA = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,11));
        sheet.setColumnWidth(0, 2600);
        sheet.setColumnWidth(1, 2150);
        sheet.setColumnWidth(2, 11300);
        sheet.setColumnWidth(3, 2400);
        for (int i = 4; i < 10; i++)
        	sheet.setColumnWidth(i, 1700);
        sheet.setColumnWidth(10, 6000);
        sheet.setColumnWidth(11, 4800);
        
        
        // Create cells for Row A
        for(int i = 0; i < 12; i++) {
            XSSFCell cell = headerRowA.createCell(i);
            cell.setCellStyle(headerRowACellStyle);
            if (i == 0)
            	cell.setCellValue(sheetName + " Schedule for " + weekStartDate + " to " + weekEndDate);
        }
        
        /*
         * 
         * Create Header Row B
         * 
         */
        // Create a Font for styling header cells
        XSSFFont headerRowBFont = workbook.createFont();
        headerRowBFont.setFontName("ARIAL");
        headerRowBFont.setFontHeightInPoints((short) 10);
        
        // Create a CellStyle with the font
        XSSFCellStyle headerRowBCellStyle = workbook.createCellStyle();
        headerRowBCellStyle.setFont(headerRowBFont);
        headerRowBCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerRowBCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerRowBCellStyle.setBorderBottom(BorderStyle.THIN);
        headerRowBCellStyle.setBorderTop(BorderStyle.THIN);
        headerRowBCellStyle.setBorderRight(BorderStyle.THIN);
        headerRowBCellStyle.setBorderLeft(BorderStyle.THIN);
        
        // Create Row B, merge
        XSSFRow headerRowB = sheet.createRow(1);
        sheet.addMergedRegion(new CellRangeAddress(1,1,3,9));
        
        // Create cells for Row B
        for(int i = 0; i < 12; i++) {
            XSSFCell cell = headerRowB.createCell(i);
            cell.setCellStyle(headerRowBCellStyle);
        }

        /*
         * 
         * Create Header Row C
         * 
         */
        // Create a Font for styling header cells
        XSSFFont headerRowCFont = workbook.createFont();
        headerRowCFont.setFontName("ARIAL");
        headerRowCFont.setFontHeightInPoints((short) 8);
        
        // Create a CellStyle with the font
        XSSFCellStyle headerRowCCellStyle = workbook.createCellStyle();
        headerRowCCellStyle.setFont(headerRowCFont);
        headerRowCCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerRowCCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerRowCCellStyle.setBorderBottom(BorderStyle.THIN);
        headerRowCCellStyle.setBorderTop(BorderStyle.THIN);
        headerRowCCellStyle.setBorderRight(BorderStyle.THIN);
        headerRowCCellStyle.setBorderLeft(BorderStyle.THIN);
        
        // Create Row C, merge
        XSSFRow headerRowC = sheet.createRow(2);
        
        // Create cells for Row C
        for(int i = 0; i < 12; i++) {
            XSSFCell cell = headerRowC.createCell(i);
            cell.setCellStyle(headerRowCCellStyle);
            switch(i) {
            case 0:
            	cell.setCellValue("Start Time");
            	break;
            case 1:
            	cell.setCellValue("End Time");
            	break;
            case 2:
            	cell.setCellValue("Element");
            	break;
            case 3:
            	cell.setCellValue("B/L");
            	break;
            case 4:
            	cell.setCellValue("CDLMS1");
            	break;
            case 5:
            	cell.setCellValue("CDLMS2");
            	break;
            case 6:
            	cell.setCellValue("UMG1");
            	break;
            case 7:
            	cell.setCellValue("UMG2");
            	break;
            case 8:
            	cell.setCellValue("CEC");
            	break;
            case 9:
            	cell.setCellValue("JMCIS");
            	break;
            case 10:
            	cell.setCellValue("Responsible Individual(s)");
            	break;
            case 11:
            	cell.setCellValue("Support");
            	break;
            }
        }	
	}
	
	public void createDateRow(String sheetName, int rowNumber, String day, String date) {
		
        // Create a Font for styling header cells
        XSSFFont dateRowFont = workbook.createFont();
        dateRowFont.setFontName("ARIAL");
        dateRowFont.setFontHeightInPoints((short) 8);
        dateRowFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        dateRowFont.setBold(true);
        
        // Create a CellStyle with the font
        XSSFCellStyle dateRowCellStyle = workbook.createCellStyle();
        dateRowCellStyle.setFont(dateRowFont);
        dateRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dateRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dateRowCellStyle.setBorderBottom(BorderStyle.THIN);
        dateRowCellStyle.setBorderTop(BorderStyle.THICK);
        dateRowCellStyle.setBorderRight(BorderStyle.THIN);
        dateRowCellStyle.setBorderLeft(BorderStyle.THIN);
        
        // Create Row B, merge
        XSSFSheet sheet = workbook.getSheet(sheetName);
        XSSFRow dateRow = sheet.createRow(rowNumber);
        
        // Create cells for Row B
        for(int i = 0; i < 12; i++) {
            XSSFCell cell = dateRow.createCell(i);
            cell.setCellStyle(dateRowCellStyle);
            switch(i) {
            case 0:
            	cell.setCellValue(day);
            	break;
            case 1:
            	cell.setCellValue(date);
            	break;
            }
        }
		
	}
	
	//TODO add resources to shot
	public void addShotToSchedule(String sheetName, int rowNumber, String shotName, String startTime, String endTime, String resources, String ri) {
		
		// Create a Font for styling new row
        XSSFFont newRowFont = workbook.createFont();
        newRowFont.setFontName("ARIAL");
        newRowFont.setFontHeightInPoints((short) 9);
        newRowFont.setBold(false);
        
        // Create a Font for styling time fields in new row
        XSSFFont newRowTimeFont = workbook.createFont();
        newRowTimeFont.setFontName("ARIAL");
        newRowTimeFont.setFontHeightInPoints((short) 8);
        newRowTimeFont.setBold(false);
        
        // Create a CellStyle with the font
        XSSFCellStyle newRowCellStyle = workbook.createCellStyle();
        newRowCellStyle.setFont(newRowFont);
        newRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newRowCellStyle.setBorderBottom(BorderStyle.THIN);
        newRowCellStyle.setBorderTop(BorderStyle.THIN);
        newRowCellStyle.setBorderRight(BorderStyle.THIN);
        newRowCellStyle.setBorderLeft(BorderStyle.THIN);
        
        // Create a CellStyle with the font for the time fields
        DataFormat format = workbook.createDataFormat(); // Sets up format for the time fields
        XSSFCellStyle newRowTimeCellStyle = workbook.createCellStyle();
        newRowTimeCellStyle.setFont(newRowTimeFont);
        newRowTimeCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newRowTimeCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newRowTimeCellStyle.setBorderBottom(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderTop(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderRight(BorderStyle.THIN);
        newRowTimeCellStyle.setBorderLeft(BorderStyle.THIN);
        newRowTimeCellStyle.setDataFormat(format.getFormat("0000"));
        
        // Add row to sheet
        XSSFSheet sheet = workbook.getSheet(sheetName);
        XSSFRow newRow = sheet.createRow(rowNumber);
        
        // Split the resources for each shot into an array
        String[] resourceArray = resources.split(",");
        
        // Create the cells for the new row
        for(int i = 0; i < 12; i++) {
        	
            XSSFCell cell = newRow.createCell(i);
            
            switch(i) {
            case 0:
            	
            	cell.setCellStyle(newRowTimeCellStyle);
            	cell.setCellValue(Integer.parseInt(startTime));
            	break;
            	
            case 1:
            	
            	cell.setCellStyle(newRowTimeCellStyle);
            	cell.setCellValue(Integer.parseInt(endTime));
            	break;
            	
            case 2:
            	
            	cell.setCellStyle(newRowCellStyle);
            	
            	String buildId = "";
            	for (String res : resourceArray) {
            		if (res.contains("BUILD:")) {
            			buildId = res.replace(" BUILD: ", "");
            			break;
            		}
            	}
            	
            	cell.setCellValue(shotName + " (" + buildId + ")");
            	break;
            	
            case 3:
            	
            	cell.setCellStyle(newRowCellStyle);
            	
            	String configId = "";
            	for (String res : resourceArray) {
            		if (res.contains("CONFIG:")) {
            			configId = res.replace(" CONFIG: ", "");
            			break;
            		}
            	}
            	
            	cell.setCellValue(configId);
            	break;
            	
            case 4:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("CDLMS1"))
            		cell.setCellValue("X");
            	
            	break;
            	
            case 5:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("CDLMS2"))           		
            		cell.setCellValue("X");
            	
            	break;
            	
            case 6:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("UMG1"))            		
            		cell.setCellValue("X");
            	
            	break;
            	
            case 7:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("UMG2"))
            		cell.setCellValue("X");
            	
            	break;
            	
            case 8:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("CEC"))
            		cell.setCellValue("X");
            	
            	break;
            	
            case 9:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("JMCIS"))
            		cell.setCellValue("X");
            	
            	break;
            	
            case 10:
            	
            	cell.setCellStyle(newRowCellStyle);
            	
            	// Split the RI input to evaluate for multiple RI's
            	String[] splitStr = ri.split(",");
            	
            	String formatedRiString = "";
            	if (splitStr.length > 2) { 

            		// More than 1 RI identified
            		// Format the RI string with separating characters
            		for (int j = 0; j < splitStr.length; j++) {
            			if (j == 0) 
            				formatedRiString = formatedRiString + splitStr[j];
            			else if (j % 2 == 0) 
            				formatedRiString = formatedRiString + " | " + splitStr[j];
            			else
            				formatedRiString = formatedRiString + ", " + splitStr[j];
            		}
            		
            		cell.setCellValue(formatedRiString);
            		
            		sheet.autoSizeColumn(i);
            		
            	}
            	else {
            		cell.setCellValue(ri);
            	}
            	

            	
            	
            	break;
            	
            case 11:
            	
            	cell.setCellStyle(newRowCellStyle);
            	if (resources.contains("CDLMS") || resources.contains("UMG"))
            		cell.setCellValue("MLST3");
            	
            	break;
            	
            default:
            	cell.setCellStyle(newRowCellStyle);
            }
        }
		
	}
	
	public void closeWorkbook(String filePath) {
		// Write the output to a file
        FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(filePath);
		} catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			workbook.write(fileOut);
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			fileOut.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}

        // Closing the workbook
        try {
			workbook.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void populateBulkUpload(String sheetName, int rowNumber, String shotName, String startTime, String endTime, String resources, String ri) {
		
		// Add row to sheet
		XSSFSheet bulkUploadSheet = workbook.getSheet("Shot_Template");
        //XSSFRow newRow = sheet.createRow(rowNumber);
        XSSFRow newRow = bulkUploadSheet.getRow(rowNumber);
        
        // Split the resources for each shot into an array
        String[] resourceArray = resources.split(",");
        
        // Create the cells for the new row
        for(int i = 0; i < 12; i++) {
        	
        	XSSFCell cell = newRow.getCell(i);
            
            switch(i) {
            
            case 0:
            	
            	// Split the RI input to evaluate for multiple RI's
            	String[] splitStr = ri.split(",");
            	
            	if (ri.contains("Jr."))
            		cell.setCellValue(splitStr[0] + ", " + splitStr[1] + ", " + splitStr[2]);
            	else
            		cell.setCellValue(splitStr[0] + ", " + splitStr[1]);
            	
            	break;
            	
            case 1:
            	
            	cell.setCellValue(shotName);
            		
            
            }
        	
        }
		
	}

}
