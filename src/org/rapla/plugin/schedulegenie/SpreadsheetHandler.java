package org.rapla.plugin.schedulegenie;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.fortuna.ical4j.model.Date;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.LinkedList;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException; 

/**
 * Handles all operations regarding the MS Excel Spreadsheets
 * @author Kevin Tooley
 * @version 1.0.0
 */
public class SpreadsheetHandler {
	
	// Declare the workbook used for the lab schedules
	public XSSFWorkbook workbook;
	public HSSFWorkbook bulkUpload;
	
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
	public SpreadsheetHandler(boolean isTest) throws FileNotFoundException, IOException {
		
		// Set unit test flag
		isUnitTest = isTest;
		
		// Create a Workbook for Lab schedules
        workbook = new XSSFWorkbook(); // new XSSFWorkbook() for generating `.xlsx` file
        
        //final String userHome = System.getProperty("user.home");
        //String filePath = userHome + "\\Documents\\ScheduleGenie_Zeta\\nscc_bulk_template.xls";
        String filePath = System.getProperty("user.dir") + System.getProperty("file.separator") + "data" + System.getProperty("file.separator") + "nscc_bulk_template.xls";
        System.out.println(filePath);
        
        // Create workbook for bulk upload
        bulkUpload = new HSSFWorkbook(new FileInputStream(filePath));
	}
	
	/**
	 * Creates the 3 header rows at the top of each sheet in the excel workbook
	 * @param sheetName
	 * @param weekStartDate
	 * @param weekEndDate
	 */
	public void createScheduleSheet(String sheetName, String weekStartDate, String weekEndDate) {

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        @SuppressWarnings("unused")
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
	
	/**
	 * This method creates a new date row for each day in the excel schedule
	 * @param sheetName
	 * @param rowNumber
	 * @param day
	 * @param date
	 */
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
	
	/**
	 * After the main function has created the lab and test shot objects, the test shot is passed to this method to get added to the excel spreadsheet.  
	 * @param sheetName
	 * @param rowNumber
	 * @param shotName
	 * @param startTime
	 * @param endTime
	 * @param resources
	 * @param ri
	 */
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
            			String tmp = res.replaceAll("\\s", "");
            			buildId = tmp.replace("BUILD:", "");
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
            			String tmp = res.replaceAll("\\s", "");
            			configId = tmp.replace("CONFIG:", "");
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
            	
	            cell.setCellValue(getShotRiString(ri));
	            		
	            sheet.autoSizeColumn(i);
            	
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
	
	/**
	 * Saves the workbooks using a FileOutputStream. The stream is created and the workbooks are written
	 * to the stream.  The Bulk Upload spreadsheet will always receive the same file name, but the 
	 * Schedule spreadsheet will allow the user to change the file for numerous revisions. 
	 * @param scheduleFilePath absolute path string
	 * @param bulkFilePath absolute path string
	 */
	public void closeWorkbook(String scheduleFilePath, String bulkFilePath) {
		// Write the output to a file
        FileOutputStream scheduleOutStream = null;
        FileOutputStream bulkOutStream = null;
		try {
			if (isUnitTest)
				scheduleOutStream = new FileOutputStream(scheduleFilePath);
			else
				scheduleOutStream = new FileOutputStream(chooseFile(scheduleFilePath));
			bulkOutStream = new FileOutputStream(bulkFilePath);
		} catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			workbook.write(scheduleOutStream);
			bulkUpload.write(bulkOutStream);
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
        try {
			scheduleOutStream.close();
			bulkOutStream.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}

        // Closing the workbook
        try {
			workbook.close();
			bulkUpload.close();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * Uses JFileChooser swing extension to open a dialog box.  Default location is set to 
	 * <user_home>\Documents\ScheduleGenie_Zeta\exports directory.
	 * @param suggestedFileName
	 * @return
	 */
	public String chooseFile(String suggestedFileName) {
		
		JFileChooser fileChooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
		        "EXCEL Spreadsheets", "xlsx", "xls");
		fileChooser.setFileFilter(filter);
		fileChooser.setCurrentDirectory(new File
				(System.getProperty("user.home") + System.getProperty("file.separator") + "Documents" + System.getProperty("file.separator") + "ScheduleGenie_Zeta" + System.getProperty("file.separator") + "exports"));
		fileChooser.setSelectedFile(new File (suggestedFileName));
		//if (fileChooser.showOpenDialog(fileChooser) == JFileChooser.APPROVE_OPTION) {
		//if (fileChooser.showSaveDialog(fileChooser) == JFileChooser.APPROVE_OPTION) {
		if (fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
		  File file = fileChooser.getSelectedFile();
		  //System.out.println(file.getName());
		  //System.out.println(file.getAbsolutePath());
		  return file.getAbsolutePath();
		}
		return "failed";
		
	}
	
	/**
	 * After the main function creates the objects and parses the database, each shot is sent to this method.
	 * This method simply loops through each cell in a given row and assigns the given values to the cell.
	 * @param sheetName hard coded to "Shot_Template"
	 * @param labName
	 * @param rowNumber
	 * @param shotName
	 * @param date
	 * @param startTime
	 * @param endTime
	 * @param resources
	 * @param ri
	 */
	public void populateBulkUpload(String sheetName, String labName, int rowNumber, String shotName, String date, String startTime, String endTime, String resources, String ri) {
		
		// Get the lab configuration from the config file
		InputHandler labConfig = new InputHandler();
		labConfig.parseCsv(new File(System.getProperty("user.dir") + System.getProperty("file.separator") + "data" + System.getProperty("file.separator") + "lab_configuration.cfg"));
		labConfig.parseCsv(new File(System.getProperty("user.dir") + System.getProperty("file.separator") + "data" + System.getProperty("file.separator") + "configuration_mapping.cfg"));
		final LinkedList<LabMapping> LAB_MAPS = labConfig.getMapping();
		final LinkedList<ConfigMapping> CONFIG_MAPS = labConfig.getConfigMapping();
		
		// Add row to sheet
		HSSFSheet bulkUploadSheet = bulkUpload.getSheet(sheetName);
        //XSSFRow newRow = sheet.createRow(rowNumber);
        //HSSFRow newRow = bulkUploadSheet.getRow(rowNumber);  //.createRow(rowNumber);
        HSSFRow newRow = bulkUploadSheet.createRow(rowNumber);
        
        // Split the resources for each shot into an array
        String[] resourceArray = resources.split(",");
        
        // Create the cells for the new row
        for(int i = 0; i < 18; i++) {
        	
        	//HSSFCell cell = newRow.getCell(i);
        	HSSFCell cell = newRow.createCell(i);
            
        	String cellValue = "";  // used to populate cell after switch statement below
        	
            switch(i) {
            
            case 0:
            	
            	// Split the RI input to evaluate for multiple RI's
            	String[] splitStr = ri.split(",");
            	
            	//This handles the case if there is a suffix in the name (i.e. Jr., Sr., etc)
            	if (splitStr[1].contains("Jr.") || splitStr[1].contains("Sr.")) 
            		cellValue = splitStr[0] + "," + splitStr[1] + "," + splitStr[2];
            	else 
            		cellValue = splitStr[0] + "," + splitStr[1];
            	
            	break;
            	
            case 1:
            	
            	// Sets the "Amplifying Information" field in the bulk upload template
            	//cell.setCellValue(shotName);
            	cellValue = shotName;
            	break;
            	
            case 2:
            	break;
            case 3:
            	
            	/*
            	 * Look in each item of the resourceArray.
            	 * If the resource is a "CONFIG:" resource, parse the item and remove whitespace
            	 * Search the CONFIG_MAP to find the appropriate mapping for the TSSS
            	 */
            	for (String res : resourceArray) {
            		if (res.contains("CONFIG:")) {
            			String tmp = res.replaceAll("\\s", "");
            			String config = tmp.replace("CONFIG:", "");
            			
            			for (ConfigMapping configMap : CONFIG_MAPS) {
                    		if (configMap.getCommon_config_name().equals(config))
                    			cellValue = configMap.getTsss_config_name();
                    	}
            			break;
            		}
            	}
            	break;
            case 4:
            	for (String res : resourceArray) {
            		if (res.contains("TE:")) {
            			String tmp = res.replaceAll("\\s", "");
            			cellValue = tmp.replace("TE:", "");
            			break;
            		}
            	}
            	break;
            	
            case 5:
            	for (String res : resourceArray) {
            		if (res.contains("ELEMENT:")) {
            			String tmp = res.replaceAll("\\s", "");
            			cellValue = tmp.replace("ELEMENT:", "");
            			break;
            		}
            	}
            	break;
            case 6:

            	//SimpleDateFormat datetemp = new SimpleDateFormat("MM/d/yyyy");
            	//Date cellDateValue = datetemp.parse(date);
            	//cellValue = 
            	
            	cellValue = date;
            	break;
            	
            case 7:

            	cellValue = startTime;
            	break;
            	
            case 8:

            	cellValue = endTime;
            	break;
            	 	
            case 9:
            	
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name1();
            	}

            	break;
            	
            case 10:
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name2();
            	}

            	break;
            case 11:
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name3();
            	}

            	break;
            case 12:
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name4();
            	}

            	break;
            case 13:
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name5();
            	}

            	break;
            case 14:
            	for (LabMapping lab : LAB_MAPS) {
            		if (lab.getCommon_name().equals(labName))
            			cellValue = lab.getTsss_name6();
            	}

            	break;
            case 15:
            	for (String res : resourceArray) {
            		if (res.contains("CDLMS") || res.contains("UMG")) {
            			cellValue = res.replaceAll("\\s", "");
            			break;
            		}
            	}
            	break;
            case 16:
            	for (String res : resourceArray) {
            		if (res.contains("CDLMS")) {
            			cellValue = "MLST3 (" + res.replaceAll("\\s", "") + ")";
            			break;
            		}
            		else if (res.contains("UMG1")) {
            			cellValue = "UMG-1 SUPPORT";
            			break;
            		}
            		else if (res.contains("UMG2")) {
            			cellValue = "UMG-2 SUPPORT";
            			break;
            		}
            	}
            	break;
            case 17:
            	for (String res : resourceArray) {
            		if (res.contains("LIVE CEC")) {
            			cellValue = "LIVE CEC/WASP";
            			break;
            		}
            	}
            	break;
            	
            }
            
            // If if statement added to format the output properly for the start and end time
            if (i == 7 || i == 8)
            	cell.setCellValue(Integer.parseInt(cellValue));
            else if (i == 6) {
            	
            	// Added the following to format the bulk upload spreadsheet "Date" field
            	// An exception was thrown every time the operator attempted to upload the spreadsheet
            	// Exception called out a type mismatch
            	// Investigation revealed that the field switched to a CellType.STRING.  The following was added
            	// to force the cell to maintain CellType.NUMERIC with a "Date" format
            	java.util.Date datetemp = null;
            	SimpleDateFormat format = new SimpleDateFormat("M/d/yy");
            	try {
					datetemp = format.parse(cellValue);
				} catch (ParseException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            	
            	cell.setCellType(CellType.NUMERIC);
            	HSSFCellStyle style = bulkUpload.createCellStyle();
            	style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy"));
            	cell.setCellValue(datetemp);
            	cell.setCellStyle(style);
            }
            else
            	cell.setCellValue(cellValue);
        	
        }
		
	}
	
	/**
	 * Called from case 10 of the addShotToSchedule method.  This returns a formated string of the shot owners
	 * @param ri string from rapla ri field
	 * @return string (formatted)
	 */
	public String getShotRiString(String ri) {
		
		String formatedString = "";
		
		// Split the input ri string into an array and strip the whitespace
		String[] splitStr = ri.replace(" ", "").split(",");
		
		final int arrayLength = splitStr.length;
		int suffixCount = 0, fromIndex = 0;
        
		// Count the number of times a name suffix (i.e. "Jr.") is in the ri string
        while ((fromIndex = ri.indexOf(".", fromIndex)) != -1 ){
            suffixCount++;
            fromIndex++;
        }
		
        /* 
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
        */
		switch (arrayLength) {
		
		case 2:
			formatedString = splitStr[1] + " " + splitStr[0];
			break;
			
		case 3:
			formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1];
			break;
			
		case 4:
			formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[3] + " " + splitStr[2];
			break;
			
		case 5:
			if (splitStr[1].equals("Jr.") || splitStr[1].equals("Sr."))
				formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1] + " | " + splitStr[4] + " " + splitStr[3];
			else
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[4] + " " + splitStr[2] + ", " + splitStr[3];
			break;
			
		case 6:
			if (suffixCount == 0)
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[3] + " " + splitStr[2] + " | " + splitStr[5] + " " + splitStr[4];
			else
				formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1] + " | " + splitStr[5] + " " + splitStr[3] + ", " + splitStr[4];
			break;
			
		case 7:
			if (splitStr[1].equals("Jr.") || splitStr[1].equals("Sr."))
				formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1] + " | " + splitStr[4] + " " + splitStr[3] + " | " + splitStr[6] + " " + splitStr[5];
			else if (splitStr[3].equals("Jr.") || splitStr[3].equals("Sr."))
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[4] + " " + splitStr[2] + ", " + splitStr[3] + " | " + splitStr[6] + " " + splitStr[5];
			else if (splitStr[5].equals("Jr.") || splitStr[5].equals("Sr."))
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[3] + " " + splitStr[2] + " | " + splitStr[6] + " " + splitStr[4] + ", " + splitStr[5];
			break;
			
		case 8:
			if (suffixCount == 0)
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[3] + " " + splitStr[2] + " | " + splitStr[5] + " " + splitStr[4] + " | " + splitStr[7] + " " + splitStr[6];
			else if (splitStr[1].equals("Jr.") || splitStr[1].equals("Sr."))
				if (splitStr[4].equals("Jr.") || splitStr[4].equals("Sr."))
					formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1] + " | " + splitStr[5] + " " + splitStr[3] + ", " + splitStr[4] + " | " + splitStr[7] + " " + splitStr[6];
				else
					formatedString = splitStr[2] + " " + splitStr[0] + ", " + splitStr[1] + " | " + splitStr[4] + " " + splitStr[3] + " | " + splitStr[7] + " " + splitStr[5] + ", " + splitStr[6];
			else
				formatedString = splitStr[1] + " " + splitStr[0] + " | " + splitStr[4] + " " + splitStr[2] + ", " + splitStr[3] + " | " + splitStr[7] + " " + splitStr[5] + ", " + splitStr[6];
			break;
		case 9:
		case 10:
		case 11:
		case 12:
			formatedString = "NAME ERROR: EXCEEDED THE NUMBER OF SHOT OWNERS";
			break;
		default:
			formatedString = "NAME ERROR: CHECK INPUTS";
		
		}
		
		return formatedString;
	}

}
