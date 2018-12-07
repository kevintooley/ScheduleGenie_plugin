package org.rapla.plugin.schedulegenie;

import java.awt.Component;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import javax.swing.JMenuItem;
import javax.swing.SwingUtilities;

import org.rapla.RaplaMainContainer;
import org.rapla.components.iolayer.IOInterface;
import org.rapla.entities.User;
import org.rapla.entities.domain.AppointmentBlock;
import org.rapla.facade.CalendarSelectionModel;
import org.rapla.framework.RaplaContext;
import org.rapla.framework.RaplaException;
import org.rapla.gui.RaplaGUIComponent;
import org.rapla.gui.toolkit.DialogUI;
import org.rapla.gui.toolkit.IdentifiableMenuEntry;
import org.rapla.plugin.tableview.RaplaTableColumn;
import org.rapla.plugin.tableview.TableViewExtensionPoints;
import org.rapla.plugin.tableview.internal.ReservationTableViewFactory;
import org.rapla.plugin.tableview.internal.TableConfig;

public class ScheduleGenieMenu extends RaplaGUIComponent implements IdentifiableMenuEntry, ActionListener {
	
	String id = "schedule.export";
	JMenuItem item;

	public ScheduleGenieMenu(RaplaContext sm) {
		super(sm);
		setChildBundleName(ScheduleGeniePlugin.RESOURCE_FILE);
		item = new JMenuItem(getString(id));
		item.setIcon(getIcon("icon.export"));
		item.addActionListener(this);
	}
	
	 public void actionPerformed(ActionEvent evt) {
         try {
         	CalendarSelectionModel model = getService(CalendarSelectionModel.class);
             export( model);
         } catch (Exception ex) {
             showException( ex, getMainComponent() );
         }
     }
	
	public String getId() {
		//return idString;
		return id;
	}

	public JMenuItem getMenuElement() {
		//return exportEntry;
		return item;
	}
	
	private static final String LINE_BREAK = "\n"; 
	private static final String CELL_BREAK = ";"; 
	
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void export(final CalendarSelectionModel model) throws Exception
	{
	    // generates a text file from all filtered events;
	    StringBuffer buf = new StringBuffer();
	    
	    Collection< ? extends RaplaTableColumn<?>> columns;
	    List<Object> objects = new ArrayList<Object>();
	    User user = model.getUser();
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
	    for (RaplaTableColumn column: columns)
    	{
	    	buf.append( column.getColumnName());
	    	buf.append(CELL_BREAK);
    	}
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
				
		final String userHome = System.getProperty("user.home");
		
		String filename = calendarName + "-" + sdfyyyyMMdd.format( model.getStartDate() )  + "-" + sdfyyyyMMdd.format( model.getEndDate() ) + ".csv";
		
		/* Keeping the below filename as individual piece as this works better for the resultant SaveAs dialog box
		   By concatenating userHome\Documents\filename together, it throws off the dialog box and makes it more
		   difficult for the user */
		if (saveFile( bytes, userHome + "\\Documents\\" + filename,"csv"))
		{
			exportFinished(getMainComponent());
		}
		
		InputHandler handler = new InputHandler();
		handler.parseCsv(userHome + "\\Documents\\" + filename);  // See note above describing reason for this filename
		
	}
	
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

	private String escape(Object cell) { 
		return cell.toString().replace(LINE_BREAK, " ").replace(CELL_BREAK, " "); 
	}
	
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

	
}
