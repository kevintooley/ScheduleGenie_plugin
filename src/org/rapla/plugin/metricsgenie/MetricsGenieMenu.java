package org.rapla.plugin.metricsgenie;

import java.awt.Component;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;

import javax.swing.JFrame;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;

import org.rapla.components.iolayer.IOInterface;
import org.rapla.entities.User;
import org.rapla.entities.domain.Allocatable;
import org.rapla.entities.domain.AppointmentBlock;
import org.rapla.facade.CalendarSelectionModel;
import org.rapla.framework.RaplaContext;
import org.rapla.framework.RaplaException;
import org.rapla.framework.logger.Logger;
import org.rapla.gui.RaplaGUIComponent;
import org.rapla.gui.toolkit.DialogUI;
import org.rapla.gui.toolkit.IdentifiableMenuEntry;
import org.rapla.plugin.metricsgenie.MetricsGeniePlugin;
import org.rapla.plugin.schedulegenie.Lab;
import org.rapla.plugin.schedulegenie.SpreadsheetHandler;
import org.rapla.plugin.tableview.RaplaTableColumn;
import org.rapla.plugin.tableview.TableViewExtensionPoints;
import org.rapla.plugin.tableview.internal.TableConfig;

import javafx.scene.Scene;
import javafx.scene.layout.VBox;

public class MetricsGenieMenu extends RaplaGUIComponent implements IdentifiableMenuEntry, ActionListener {
	
	String id = "schedule.metrics";
	JMenuItem item;
	final Logger logger = getLogger();

	/**
	 * Setup the resources for the "Export Metrics" menu item
	 * @param context Returns a reference to the requested object (e.g. a component instance)
	 */
	public MetricsGenieMenu(RaplaContext context) {
		super(context);
		
		// Create the menu entry
		setChildBundleName(MetricsGeniePlugin.RESOURCE_FILE);
		item = new JMenuItem(getString(id));
		item.setIcon(getIcon("icon.export"));
		item.addActionListener(this);
		
		logger.info("MetricsGenieMenu plugin started.");
	}
	
	/**
	 * Obtain the calendar data from the current week (visible on screen) and export the data to a file and to the InputHandler class
	 */
	public void actionPerformed(ActionEvent evt) {
		try {
		 	CalendarSelectionModel model = getService(CalendarSelectionModel.class);
		    export( model);
		    logger.info("Exported Model");
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
		 */
		Calendar startCalendar = Calendar.getInstance();
		Calendar stopCalendar = Calendar.getInstance();
		int M, D, Y;
		
		String[] startDateString = JOptionPane.showInputDialog("Enter a start date (mm/dd/yy)").split("/");
		if (	((Integer.parseInt(startDateString[0]) > 0) && (Integer.parseInt(startDateString[0]) <= 12)) &&
				((Integer.parseInt(startDateString[1]) > 0) && (Integer.parseInt(startDateString[1]) <= 31)) &&
				((Integer.parseInt(startDateString[2]) >= 0) && (Integer.parseInt(startDateString[2]) <= 99))	) {
			M = Integer.parseInt(startDateString[0]) - 1;  // the Calendar API seems to be a zero-based enum; therefore subtract 1
			D = Integer.parseInt(startDateString[1]);
			Y = Integer.parseInt(startDateString[2]) + 2000;
			startCalendar.set(Y, M, D, 0, 0, 0);
		}
		
		//JOptionPane.showMessageDialog(null, "Start date: " + startCalendar.getTime());
		
		String[] stopDateString = JOptionPane.showInputDialog("Enter a stop date (mm/dd/yy)").split("/");
		if (	((Integer.parseInt(stopDateString[0]) > 0) && (Integer.parseInt(stopDateString[0]) <= 12)) &&
				((Integer.parseInt(stopDateString[1]) > 0) && (Integer.parseInt(stopDateString[1]) <= 31)) &&
				((Integer.parseInt(stopDateString[2]) >= 0) && (Integer.parseInt(stopDateString[2]) <= 99))	) {
			M = Integer.parseInt(stopDateString[0]) - 1;
			//D = Integer.parseInt(stopDateString[1]) + 1;  // Adding 1 day in order to cover all shots (up to midnight) on the date entered
			D = Integer.parseInt(stopDateString[1]);
			Y = Integer.parseInt(stopDateString[2]) + 2000;
			stopCalendar.set(Y, M, D, 23, 59, 59);
		}
		
		//JOptionPane.showMessageDialog(null, "Stop date: " + stopCalendar.getTime());
		
		Date newStart = startCalendar.getTime();
		model.setStartDate(newStart);
		
		Date newEnd = stopCalendar.getTime();
		model.setEndDate(newEnd);
		
		logger.info("Start Date: " + model.getStartDate());
		logger.info("End Date: " + model.getEndDate());
		
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
		 * Create a Schedule Sheet for each lab
		 */
		MetricsSpreadsheetHandler sh = new MetricsSpreadsheetHandler(false);
		sh.createScheduleSheet();
		
		// Row counter for bulk upload spreadsheet starts at row 1; 
		// Isolated from loop below as we don't want to reset this counter
		int j = 1;
		
		logger.info("Rapla TimeZone: " + getRaplaLocale().getTimeZone());
		
		
		/*
		 * Create the lab object for each lab
		 */
		for (Allocatable lab : model.getSelectedAllocatables()) {
			
			String name = lab.getName(getLocale());
			labList.add(name);
			
			/*
			 * Search for shots in each lab
			 */
			logger.info("Searching for " + name + " shots...");
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
		
		for (Lab room : labObjects) {
			
			for (Object app : room.shots) {
				
				ArrayList<Object> shotData = new ArrayList<Object>();
				
				for (RaplaTableColumn column : myColumns) {
					
					Object value = column.getValue(app);
					
					Class columnClass = column.getColumnClass();
		    		boolean isDate = columnClass.isAssignableFrom( java.util.Date.class);
					
		    		if (value != null) {
		    			if ( isDate) { 
							
		    				Date tempDate = (Date) value;
		    				
		    				SimpleDateFormat format = new SimpleDateFormat();
		    				format.setTimeZone(TimeZone.getTimeZone("EDT"));
		    				
		    				shotData.add(format.format(tempDate));
		    				
		    				logger.info("Time: " + format.format(tempDate));
							
						} else {
							
							shotData.add(value);
							logger.info((String)value);
							
						}
		    			
		    		}
	
				}
				
				sh.addShotToSchedule(room.name, j, shotData);
				logger.info("----- " + j + " ------");
				j++;
				
			}
			
		}	
		
		//sh.closeWorkbook("C:/Users/ktooley/Documents/TEST/metric-FullOutput.xlsx");  //FOR TEST PURPOSES
		
		DateFormat sdfyyMMdd = new SimpleDateFormat("yyMMdd");
		sdfyyMMdd.setTimeZone(TimeZone.getTimeZone("GMT"));
		logger.info("Setting timestamp for filename...");
		
		// Use a simple string for the filename instead of the long sequence commented below
		final String scheduleName = "_Test_Schedule_Metrics";
		
		final String scheduleFilename = System.getProperty("user.dir") + System.getProperty("file.separator") + "exports" + System.getProperty("file.separator") + sdfyyMMdd.format( model.getStartDate() ) + scheduleName + ".xlsx";
		
		logger.info("Saving file " + scheduleFilename);
		sh.closeWorkbook(scheduleFilename);
		logger.info("File saved.  Mission complete.");

	}
	

}
