package org.rapla.plugin.schedulegenie;

import java.awt.Component;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;

import javax.swing.JMenuItem;
import javax.swing.SwingUtilities;

import org.rapla.components.iolayer.IOInterface;
import org.rapla.entities.User;
import org.rapla.entities.domain.Allocatable;
import org.rapla.entities.domain.AppointmentBlock;
import org.rapla.facade.CalendarSelectionModel;
import org.rapla.framework.RaplaContext;
import org.rapla.framework.RaplaException;
import org.rapla.gui.RaplaGUIComponent;
import org.rapla.gui.toolkit.DialogUI;
import org.rapla.gui.toolkit.IdentifiableMenuEntry;
import org.rapla.plugin.tableview.RaplaTableColumn;
import org.rapla.plugin.tableview.TableViewExtensionPoints;
import org.rapla.plugin.tableview.internal.TableConfig;

/**
 * The ScheduleGenieMenu class adds the menu and passes the calendar view data to the InputHandler class for processing
 * @author Kevin Tooley
 * @version 1.0.0
 */
public class ScheduleGenieMenu extends RaplaGUIComponent implements IdentifiableMenuEntry, ActionListener {
	
	String id = "schedule.export";
	JMenuItem item;

	/**
	 * Setup the resources for the "Export To ScheduleGenie" menu item
	 * @param context Returns a reference to the requested object (e.g. a component instance)
	 */
	public ScheduleGenieMenu(RaplaContext context) {
		super(context);
		
		// Create the menu entry
		setChildBundleName(ScheduleGeniePlugin.RESOURCE_FILE);
		item = new JMenuItem(getString(id));
		item.setIcon(getIcon("icon.export"));
		item.addActionListener(this);
	}
	
	/**
	 * Obtain the calendar data from the current week (visible on screen) and export the data to a file and to the InputHandler class
	 */
	public void actionPerformed(ActionEvent evt) {
		try {
		 	CalendarSelectionModel model = getService(CalendarSelectionModel.class);
		    export( model);
		} catch (Exception ex) {
		    showException( ex, getMainComponent() );
		}
	}
	
	public String getId() {
		return id;
	}

	public JMenuItem getMenuElement() {
		return item;
	}
	
	private static final String LINE_BREAK = "\n"; 
	private static final String CELL_BREAK = ";"; 
	
	/**
	 * Export the data from the current visible week to csv file and the InputHandler class
	 * @param model is the object representing the Rapla data
	 * @throws Exception
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void export(final CalendarSelectionModel model) throws Exception
	{
				
		/*
		 * Preconditions and Setup
		 */
		
		// Create user; needed for LoadColumns API
		User myUser = model.getUser();
		
		/*
		 * Rapla primarily uses GMT as the standard for dates within the program.  However
		 * these dates are translated to local time when displayed to the operator.  Because
		 * of the use of GMT, we have to initialize things is the appropriate manner in order
		 * to get dates that mean what we want them to mean.  Below I used the calendar
		 * to manipulate the various times, eventually setting the needed Date object with
		 * the calendar output.  The Calendar API is an easy and efficient method to do this.
		 * 
		 * First, set the date in the model based on the current view in Rapla
		 */
		Calendar cal = Calendar.getInstance();
		cal.setTime(model.getSelectedDate());
		
		// Set the time to Monday at midnight of the given week
		cal.setTimeZone(TimeZone.getTimeZone("GMT"));
		cal.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY);
		cal.set(Calendar.HOUR, 0);
		cal.set(Calendar.AM_PM, Calendar.AM);
		cal.set(Calendar.MINUTE, 0);
		cal.set(Calendar.SECOND, 0);		
		
		/*
		 * The newStart and newEnd are used for appointment calculations later in this function.
		 * newStart is set to Monday at 0000.  newEnd adds 7 days to this (i.e. next Monday
		 * at 0000).  This allows for a shot that ends at 2400 on Sunday night.  Both the newStart
		 * and newEnd are used to set the model start and end points.
		 */
		Date newStart = cal.getTime();
		model.setStartDate(newStart);
		
		Date newEnd = getDate(newStart, false);
		model.setEndDate(newEnd);
		
		/*
		 * The model.getStartDate().toLocaleString() method is not returning the correct time stamp.  It 
		 * seems to be returning GMT.  As a result, I needed to use the below calculations/methods in order
		 * to produce startDate and endDate.  
		 * FIXME: fix the toLocaleString method
		 * 
		 * The following sets the date on the top of the spreadsheets.  These dates do not affect the
		 * dates of the appointment objects.  Those are handled above in the model.
		 */
		Date startDate = model.getStartDate();
		
		// Reset the calendar object to the model view (i.e. the week that Rapla is displaying)
		cal.setTime(model.getSelectedDate());
		
		// Set the current time to the following Sunday at 2359
		cal.add(Calendar.DAY_OF_MONTH, 6);
		cal.add(Calendar.MINUTE, -1);
		
		/* 
		 * Set the endDate to the cal value.  This will be displayed in the header of each schedule.
		 * endDate has no other roles or needs
		 */
		Date endDate = cal.getTime();
		
		// List of labs
		final ArrayList<String> labList = new ArrayList<String>();
		ArrayList<Lab> labObjects = new ArrayList<Lab>();
		
		// Create a collection of columns 
		Collection< ? extends RaplaTableColumn<?>> myColumns;
		myColumns = TableConfig.loadColumns(getContainer(),"appointments",TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, myUser);
		
		// Add all appointment blocks to a list of objects
		List<Object> myObjects = new ArrayList<Object>();
		final List<AppointmentBlock> myBlocks = model.getBlocks();
		
		// Add appointments to the myObjects list, but only add appointments for the current week
		// Deprecated after addition start and end time modifications above
		/*for (AppointmentBlock block : myBlocks) {
			if (block.getAppointment().getStart().before(endDate))
				myObjects.add(block);
		}*/
		
		myObjects.addAll(myBlocks);
		
		/*
		 * Create the lab object for each lab
		 */
		for (Allocatable lab : model.getSelectedAllocatables()) {
			
			String name = lab.getName(getLocale());
			labList.add(name);
			
			/*
			 * Search for shots in each lab
			 */
			// System.out.println("Searching for " + name + " shots...");
			ArrayList<Object> appointmentObjects = new ArrayList<Object>();
			
			for (Object app : myObjects) {
				
				for (RaplaTableColumn column : myColumns) {
					
					Object value = column.getValue(app);
					
					if (value.toString().contains(name)) {
						//System.out.println("value:" + value);
						//System.out.println("Shot belongs to " + name);
						appointmentObjects.add(app);
						break;
					}

				}
				
			}
			
			Lab lab1 = new Lab(name, appointmentObjects);
			labObjects.add(lab1);
			
		}
		
		System.out.println("");
		
		/*
		 * Create a Schedule Sheet for each lab
		 */
		SpreadsheetHandler sh = new SpreadsheetHandler(false);
		
		// Row counter for bulk upload spreadsheet starts at row 1; 
		// Isolated from loop below as we don't want to reset this counter
		int j = 1;  
		
		for (Lab room : labObjects) {
						
			sh.createScheduleSheet(room.name, formatShortDate(startDate), formatShortDate(endDate));
			
			//if (room.shots.size() > 0)
			//	System.out.println("Creating schedule for " + room.name);
			
			int i = 3;  // row counter starts at row 3 in Schedule; this counter resets for each sheet
			
			
			// Setup a string to track the day of week
			Date scheduleDay = startDate;
			String stringDay = formatLongDate(startDate);
			boolean incrementDay = true;  // Set to true initially to force a date line to be entered
			
			for (Object appointment : room.shots) {
				
				ArrayList<String> rowFields = new ArrayList<String>();
				
				SimpleDateFormat format = new SimpleDateFormat("HHmm");
				format.setTimeZone( getRaplaLocale().getTimeZone());
				
				//System.out.println("");
				//System.out.println("<<<BEGIN SHOT DATA>>>");
				
				for (RaplaTableColumn column : myColumns) {
					
					Object value = column.getValue(appointment);
					Class columnClass = column.getColumnClass();
		    		boolean isDate = columnClass.isAssignableFrom( java.util.Date.class);
		    		String formated = "";
					
		    		if(value != null) {
						if ( isDate)
						{ 
							/*
							 * If the value is a Date type, save a temporary string in long
							 * date format (MM/dd/yyyy).  Evaluate this string against the previously
							 * set stringDay.  If equal, the shot is in the current scheduleDay; 
							 * otherwise, increment to the next day
							 */
							String tempDate = formatLongDate( (java.util.Date)value );
							if ( !tempDate.equals(stringDay) && column.getColumnName().equals("Start")) {
								incrementDay = true;
								
								/* 
								 * 3/11/2019: Fixed bug that prevented date from rolling to next day.
								 * Issue is with local time zone.  By passing the date "value" to the getDate 
								 * method as a "startTime" (see method arguments) the return is a valid
								 */
								scheduleDay = getDate( (java.util.Date)value, true);
								stringDay = tempDate;
							}
							
							// Get and store the timestamp within the "value" Date object
							String timestamp = format.format(   (java.util.Date)value);
							formated = timestamp;
							
							// 3/11/2019: Fix bug that shows shot ending at midnight as timestamp of 0000 instead of 2400
							if ( formated.equals("0000") && column.getColumnName().equals("End") ) {
								formated = "2400";
							}
						}
						else
						{
							String escaped = escape(value);
							formated = escaped;
						}
		    		}
		    		
					rowFields.add(formated);
					
				}
				
				if (incrementDay) {
					sh.createDateRow(room.name, i, getDayOfWeek(scheduleDay), stringDay);
					incrementDay = false;
					i++;
				}
				sh.addShotToSchedule(room.name, i, rowFields.get(0), rowFields.get(1), rowFields.get(2), rowFields.get(3), rowFields.get(4));
				sh.populateBulkUpload("Shot_Template", room.name, j, rowFields.get(0), stringDay, rowFields.get(1), rowFields.get(2), rowFields.get(3), rowFields.get(4));
				
				i++;
				j++;
			}
			
		}
		
		DateFormat sdfyyMMdd = new SimpleDateFormat("yyMMdd");
		sdfyyMMdd.setTimeZone(TimeZone.getTimeZone("GMT"));
		
		// Use a simple string for the filename instead of the long sequence commented below
		final String scheduleName = "_NSCC_Test_Schedules";
		final String bulkUploadName = "_NSCC_Bulk_Upload";
		
		// Get user home property
		final String userHome = System.getProperty("user.home");
		
		String scheduleFilename = userHome + "\\Documents\\ScheduleGenie_Zeta\\exports\\" + sdfyyMMdd.format( model.getStartDate() ) + scheduleName + ".xlsx";
		String bulkFilename = userHome + "\\Documents\\ScheduleGenie_Zeta\\exports\\" + sdfyyMMdd.format( model.getStartDate() ) + bulkUploadName + ".xls";
		
		
		sh.closeWorkbook(scheduleFilename, bulkFilename);
		
		exportFinished(getMainComponent());
		
	}

		
	/**
	 * Dialog for export completion
	 * @param topLevel
	 * @return boolean
	 */
	protected boolean exportFinished(Component topLevel) {
		try {
			DialogUI dlg = DialogUI.create(
	                		 getContext()
	                		,topLevel
	                        ,true
	                        ,getString("export")
	                        ,getString("file_saved")
	                        ,new String[] { getString("ok")}
	                        );
			dlg.setIcon(getIcon("icon.export"));
	        dlg.setDefault(0);
	        dlg.start();
	        return (dlg.getSelectedIndex() == 0);
		} catch (RaplaException e) {
			return true;
		}
	
	}
	
	/**
	 * Escape method; cell replacements.
	 * @param cell
	 * @return
	 */
	private String escape(Object cell) { 
		return cell.toString().replace(LINE_BREAK, " ").replace(CELL_BREAK, " "); 
	}
	
	/**
	 * SaveAs a csv file.  Opens a dialog box
	 * @param content size of the data file
	 * @param filename sets the recommended filename
	 * @param extension is recommended as csv
	 * @return boolean
	 * @throws RaplaException
	 */
	public boolean saveFile(byte[] content, String filename, String extension) throws RaplaException {
		final Frame frame = (Frame) SwingUtilities.getRoot(getMainComponent());
		IOInterface io =  getService( IOInterface.class);
		try 
		{
			String file = io.saveFile( frame, null, new String[] {extension}, filename, content);
			return file != null;
		} 
		catch (IOException e) 
		{
			throw new RaplaException(e.getMessage(), e);
	    }
	}
	
	/**
	 * The getDate function is used in various areas of the application to return a properly
	 * formated date field to the caller.  As of March 2019, the calcStartTime portion of this 
	 * function is deprecated.  
	 * @param dateFromModel
	 * @param calcStartTime
	 * @return
	 */
	private Date getDate (Date dateFromModel, boolean calcStartTime) {
		
		Calendar cal = Calendar.getInstance();
		cal.setTime(dateFromModel);
		if (calcStartTime) {
			cal.add(Calendar.HOUR_OF_DAY, 12); // Add 12 hours to ensure the start day is monday
			cal.set(Calendar.HOUR, 0);
		}
		else {
			cal.add(Calendar.DAY_OF_MONTH, 7);
		}
		return cal.getTime();
		
	}
	
	private String formatShortDate (Date date) {
		SimpleDateFormat format = new SimpleDateFormat("MM/dd");
		format.setTimeZone(TimeZone.getTimeZone("GMT"));
		return format.format(date);
	}
	
	private String formatLongDate (Date date) {
		SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
		
		// 3/11/2019: Fixed shots that start at midnight were not scheduled on the correct day
		//format.setTimeZone( getRaplaLocale().getTimeZone());
		format.setTimeZone(TimeZone.getTimeZone("GMT"));
		
		return format.format(date);
	}

	private String getDayOfWeek (Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cal.setTimeZone(TimeZone.getTimeZone("GMT"));
		return cal.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG, Locale.getDefault());
	}
	
}
