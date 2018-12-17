package org.rapla.plugin.schedulegenie;

//TODO Remove unused imports
import java.awt.Component;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.lang.reflect.Array;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;

import javax.swing.JMenuItem;
import javax.swing.SwingUtilities;

import org.rapla.RaplaMainContainer;
import org.rapla.components.iolayer.IOInterface;
import org.rapla.entities.User;
import org.rapla.entities.domain.Allocatable;
import org.rapla.entities.domain.Appointment;
import org.rapla.entities.domain.AppointmentBlock;
import org.rapla.entities.domain.Reservation;
import org.rapla.facade.CalendarSelectionModel;
import org.rapla.framework.RaplaContext;
import org.rapla.framework.RaplaDefaultContext;
import org.rapla.framework.RaplaException;
import org.rapla.gui.RaplaGUIComponent;
import org.rapla.gui.toolkit.DialogUI;
import org.rapla.gui.toolkit.IdentifiableMenuEntry;
import org.rapla.plugin.tableview.RaplaTableColumn;
import org.rapla.plugin.tableview.TableViewExtensionPoints;
import org.rapla.plugin.tableview.internal.ReservationTableViewFactory;
import org.rapla.plugin.tableview.internal.TableConfig;
import org.rapla.storage.xml.RaplaMainReader;

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
	@SuppressWarnings({ "unchecked", "rawtypes", "deprecation" })
	public void export(final CalendarSelectionModel model) throws Exception
	{
		/*
		 * 
		 * SANDBOX
		 * 
		 */
				
		/*
		 * Preconditions and Setup
		 */
		
		// Create user; needed for LoadColumns API
		User myUser = model.getUser();
		
		/*
		 * The model.getStartDate().toLocaleString() method is not returning the correct time stamp.  It 
		 * seems to be returning GMT.  As a result, I needed to use the below calculations/methods in order
		 * to produce startDate and endDate.  
		 * FIXME: fix the toLocaleString method
		 */
		Date startDate = getDate(model.getStartDate(), true);
		//System.out.println("Start date: " + startDate);
		
		Date endDate = getDate(startDate, false);
		//System.out.println("End date is: " + endDate);
		
		// List of labs
		final ArrayList<String> labList = new ArrayList<String>();
		ArrayList<Lab> labObjects = new ArrayList<Lab>();
		
		// Create a collection of columns 
		Collection< ? extends RaplaTableColumn<?>> myColumns;
		myColumns = TableConfig.loadColumns(getContainer(),"appointments",TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, myUser);
		
		// Add all appointment blocks to a list of objects
		List<Object> myObjects = new ArrayList<Object>();
		final List<AppointmentBlock> myBlocks = model.getBlocks();
		myObjects.addAll(myBlocks);
		
		/*
		 * Print data for each shot in master myObjects array list
		 */
		/*for (Object app : myObjects) {
			System.out.println("");
			System.out.println("<<<BEGIN SHOT DATA>>>");
			int i = 0;
			for (RaplaTableColumn column : myColumns) {
				
				Object value = column.getValue(app);
				System.out.println("Column " + i + " value: " + value);
				i++;
				
			}
			System.out.println("<<<END SHOT DATA>>>");
			
		}
		
		System.out.println("");*/
		
		/*
		 * Create the lab object for each lab
		 */
		for (Allocatable lab : model.getSelectedAllocatables()) {
			
			String name = lab.getName(getLocale());
			labList.add(name);
			
			System.out.println("Searching for " + name + " shots...");
			
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
		 * Count the number of shots for each lab
		 */
		/*for (Lab room : labObjects) {
			System.out.println("Lab: " + room.name + " has " + room.shots.size() + " shots.");

		}
		
		System.out.println("");*/
		
		/*
		 * Print the shot details for each lab
		 */
		/*for (Lab room : labObjects) {
			
			if (room.shots.size() > 0)
				System.out.println("Lab: " + room.name + " has " + room.shots.size() + " shots.");
			
			for (Object app : room.shots) {
				System.out.println("");
				System.out.println("<<<BEGIN SHOT DATA>>>");
				int i = 0;
				for (RaplaTableColumn column : myColumns) {
					
					Object value = column.getValue(app);
					System.out.println("Column " + i + " value: " + value);
					i++;
					
				}
				System.out.println("<<<END SHOT DATA>>>");
				
			}
			System.out.println("");
		}*/
		
		/*
		 * Create a Schedule Sheet for each lab
		 */
		SpreadsheetHandler sh = new SpreadsheetHandler();
		
		for (Lab room : labObjects) {
						
			sh.createScheduleSheet(room.name, formatShortDate(startDate), formatShortDate(endDate));
			
			if (room.shots.size() > 0)
				System.out.println("Creating schedule for " + room.name);
			
			int i = 3;  // row counter that starts at row 3
			
			// Setup a string to track the day of week
			Date scheduleDay = startDate;
			String stringDay = formatLongDate(startDate);
			boolean incrementDay = true;  // Set to true initially to force a date line to be entered
			
			for (Object appointment : room.shots) {
				
				ArrayList<String> rowFields = new ArrayList<String>();
				
				SimpleDateFormat format = new SimpleDateFormat("HHmm");
				format.setTimeZone( getRaplaLocale().getTimeZone());
				
				System.out.println("");
				System.out.println("<<<BEGIN SHOT DATA>>>");
				
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
							if ( !tempDate.equals(stringDay) ) {
								incrementDay = true;
								scheduleDay = (java.util.Date)value;
								stringDay = tempDate;
							}
							
							// Get and store the timestamp within the "value" Date object
							String timestamp = format.format(   (java.util.Date)value);
							formated = timestamp;
						}
						else
						{
							String escaped = escape(value);
							formated = escaped;
						}
		    		}
		    		
					rowFields.add(formated);
					System.out.println("value " + formated + " moved to list");
					
				}
				System.out.println("<<<END SHOT DATA>>>");
				
				if (incrementDay) {
					sh.createDateRow(room.name, i, getDayOfWeek(scheduleDay), stringDay);
					incrementDay = false;
					i++;
				}
				sh.addShotToSchedule(room.name, i, rowFields.get(0), rowFields.get(1), rowFields.get(2), rowFields.get(3), rowFields.get(4));
				
				i++;
			}
			System.out.println("");
			
		}
		
		sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/TEST.xlsx");  // TODO: set path to user home
		
		
		/*
		 * 
		 * SANDBOX END
		 * 
		 */
		
	    // generates a text file from all filtered events;
	    StringBuffer buf = new StringBuffer();
	    
	    // Setup collections
	    Collection< ? extends RaplaTableColumn<?>> columns;
	    List<Object> objects = new ArrayList<Object>();
	    User user = model.getUser();
	    
	    // Export conditions depending on view; likely won't be applicable due to the manner that we use Rapla
	    if (model.getViewId().equals(ReservationTableViewFactory.TABLE_VIEW))
	    {
	    	columns = TableConfig.loadColumns(getContainer(),"events",TableViewExtensionPoints.RESERVATION_TABLE_COLUMN, user);
		    objects.addAll(Arrays.asList( model.getReservations())); 
	    }
	    else
	    {
	        columns = TableConfig.loadColumns(getContainer(),"appointments",TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, user);
		    final List<AppointmentBlock> blocks = model.getBlocks();
            objects.addAll( blocks); 
	    }
	    
	    // Export column names
	    for (RaplaTableColumn column: columns)
    	{
	    	buf.append( column.getColumnName());
	    	buf.append(CELL_BREAK);
    	}
	    
	    // Export Reservations
	    for (Object row: objects)
	    {
	    	buf.append(LINE_BREAK);
	    	for (RaplaTableColumn column: columns)
	    	{
	    		Object value = column.getValue( row);
	    		Class columnClass = column.getColumnClass();
	    		boolean isDate = columnClass.isAssignableFrom( java.util.Date.class);
	    		String formated = "";
	    		if(value != null) {
					if ( isDate)
					{ 
						SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
						format.setTimeZone( getRaplaLocale().getTimeZone());
						String timestamp = format.format(   (java.util.Date)value);
						formated = timestamp;
					}
					else
					{
						String escaped = escape(value);
						formated = escaped;
					}
	    		}
				buf.append( formated );
	    	   	buf.append(CELL_BREAK);
	    	}
	    }
	    
        byte[] bytes = buf.toString().getBytes();
        
		DateFormat sdfyyyyMMdd = new SimpleDateFormat("yyyyMMdd");
		
		// Use a simple string for the filename instead of the long sequence commented below
		final String calendarName = "ScheduleGenieCSV";
		//final String calendarName = getQuery().getSystemPreferences().getEntryAsString(RaplaMainContainer.TITLE, getString("rapla.title"));
		
		// Get user home property; TODO: Find method to get MyDocuments directory
		final String userHome = System.getProperty("user.home");
		
		String filename = calendarName + "-" + sdfyyyyMMdd.format( model.getStartDate() )  + "-" + sdfyyyyMMdd.format( model.getEndDate() ) + ".csv";
		
		/* Keeping the below filename as individual piece as this works better for the resultant SaveAs dialog box
		   By concatenating userHome\Documents\filename together, it throws off the dialog box and makes it more
		   difficult for the user */
		if (saveFile( bytes, userHome + "\\Documents\\" + filename,"csv"))
		{
			exportFinished(getMainComponent());
		}
		
		// Call the InputHandler to start parsing the input CSV file
		InputHandler handler = new InputHandler();
		handler.parseCsv(userHome + "\\Documents\\" + filename);  // See note above describing reason for this filename
		
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
	
	private Date getDate (Date dateFromModel, boolean calcStartTime) {
		
		Calendar cal = Calendar.getInstance();
		cal.setTime(dateFromModel);
		if (calcStartTime) {
			cal.add(Calendar.HOUR_OF_DAY, 12); // Add 12 hours to ensure the start day is monday
			cal.set(Calendar.HOUR, 0);
		}
		else {
			cal.add(Calendar.DAY_OF_MONTH, 6);
			cal.set(Calendar.HOUR, 23);
			cal.set(Calendar.MINUTE, 59);
		}
		return cal.getTime();
		
	}
	
	private String formatShortDate (Date date) {
		SimpleDateFormat format = new SimpleDateFormat("MM/dd");
		return format.format(date);
	}
	
	private String formatLongDate (Date date) {
		SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
		return format.format(date);
	}

	private String getDayOfWeek (Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		return cal.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.LONG, Locale.getDefault());
	}
	
}
