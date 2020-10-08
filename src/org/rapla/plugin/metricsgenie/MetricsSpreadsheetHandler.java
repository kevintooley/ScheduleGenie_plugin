package org.rapla.plugin.metricsgenie;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.LinkedList;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException; 

/**
 * Handles all operations regarding the MS Excel Spreadsheets
 * @author Kevin Tooley
 * @version 1.0.0
 */
public class MetricsSpreadsheetHandler {
	
	// Declare the workbook used for the lab schedules
	public XSSFWorkbook workbook;
	public HSSFWorkbook bulkUpload;
	public XSSFWorkbook updateWorkbook;
	
	private boolean isUnitTest;
	
	
	/**
	 * Accessor method for the private isUnitTest boolean
	 * @return isUnitTest boolean
	 */
	public boolean isUnitTest() {
		return isUnitTest;
	}

	/**
	 * Constructor of the SpreadsheetHandler object
	 * @param isTest denotes a unit test
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public MetricsSpreadsheetHandler(boolean isTest) throws FileNotFoundException, IOException {
		
		// Set unit test flag
		isUnitTest = isTest;
		
		// Create a Workbook for Lab schedules
        workbook = new XSSFWorkbook(); // new XSSFWorkbook() for generating `.xlsx` file
        
        
        
	}
	
	/**
	 * Creates the 3 header rows at the top of each sheet in the excel workbook
	 * @param sheetName
	 * @param weekStartDate
	 * @param weekEndDate
	 */
	public void createScheduleSheet() {

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        //@SuppressWarnings("unused")
		//XSSFCreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        XSSFSheet sheet = workbook.createSheet("Schedule");
        
        /*
         * 
         * Create Header Row A
         * 
         */
        // Create a Font for styling header cells
        //XSSFFont headerRowAFont = workbook.createFont();
        //headerRowAFont.setBold(true);
        //headerRowAFont.setFontName("ARIAL");
        //headerRowAFont.setFontHeightInPoints((short) 10);
        ////headerFont.setColor(IndexedColors.RED.getIndex());

        // TODO: Make method for CellStyle setup
        // Create a CellStyle with the font
        //XSSFCellStyle headerRowACellStyle = workbook.createCellStyle();
        //headerRowACellStyle.setFont(headerRowAFont);
        //headerRowACellStyle.setAlignment(HorizontalAlignment.CENTER);
        //headerRowACellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //headerRowACellStyle.setBorderBottom(BorderStyle.THIN);
        //headerRowACellStyle.setBorderTop(BorderStyle.THIN);
        //headerRowACellStyle.setBorderRight(BorderStyle.THIN);
        //headerRowACellStyle.setBorderLeft(BorderStyle.THIN);

        // Create Row A, merge, adjust column widths
        XSSFRow headerRowA = sheet.createRow(0);
        sheet.setColumnWidth(0, 3900);
        sheet.setColumnWidth(1, 4100);
        sheet.setColumnWidth(2, 7500);
        sheet.setColumnWidth(3, 3800);
        sheet.setColumnWidth(4, 5700);
        sheet.setColumnWidth(5, 4200);
        sheet.setColumnWidth(6, 3800);
        sheet.setColumnWidth(7, 3800);
        sheet.setColumnWidth(8, 5100);
        sheet.setColumnWidth(9, 4200);
        
        for(int i = 0; i < 10; i++) {
            XSSFCell cell = headerRowA.createCell(i);
            switch(i) {
            case 0:
            	cell.setCellValue("ELEMENT");
            	break;
            case 1:
            	cell.setCellValue("PROGRAM");
            	break;
            case 2:
            	cell.setCellValue("FUNDING SOURCE");
            	break;
            case 3:
            	cell.setCellValue("BUILD");
            	break;
            case 4:
            	cell.setCellValue("EFFORT");
            	break;
            case 5:
            	cell.setCellValue("SYSTEM");
            	break;
            case 6:
            	cell.setCellValue("START DATE");
            	break;
            case 7:
            	cell.setCellValue("END DATE");
            	break;
            case 8:
            	cell.setCellValue("TOTAL DURATION");
            	break;
            case 9:
            	cell.setCellValue("USER");
            	break;
            }
        }	
	}
	
	/**
	 * Saves the workbooks using a FileOutputStream. The stream is created and the workbooks are written
	 * to the stream.  The Bulk Upload spreadsheet will always receive the same file name, but the 
	 * Schedule spreadsheet will allow the user to change the file for numerous revisions. 
	 * @param scheduleFilePath absolute path string
	 * @param bulkFilePath absolute path string
	 */
	public void closeWorkbook(String itsFilePath) {
		// Write the output to a file
        FileOutputStream scheduleOutStream = null;

		try {
			if (isUnitTest)
				scheduleOutStream = new FileOutputStream(itsFilePath);
			else
				scheduleOutStream = new FileOutputStream(chooseFile(itsFilePath));

		} catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			workbook.write(scheduleOutStream);

		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			scheduleOutStream.close();

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
	
	/**
	 * Uses JFileChooser swing extension to open a dialog box.  Default location is set to 
	 * <user_home>\Documents\ScheduleGenie_<release>\exports directory.
	 * @param suggestedFileName
	 * @return
	 */
	public String chooseFile(String suggestedFileName) {
		
		JFileChooser fileChooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
		        "EXCEL Spreadsheets", "xlsx", "xls");
		fileChooser.setFileFilter(filter);
		
		fileChooser.setCurrentDirectory(new File
				(System.getProperty("user.dir") + System.getProperty("file.separator") + "exports"));
		
		fileChooser.setSelectedFile(new File (suggestedFileName));

		if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
		  File file = fileChooser.getSelectedFile();

		  return file.getAbsolutePath();
		}
		return "failed";
		
	}

}
