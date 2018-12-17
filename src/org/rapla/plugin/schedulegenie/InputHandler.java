package org.rapla.plugin.schedulegenie;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedList;

import org.supercsv.cellprocessor.Optional;
import org.supercsv.cellprocessor.ParseInt;
import org.supercsv.cellprocessor.ParseLong;
import org.supercsv.cellprocessor.constraint.NotNull;
import org.supercsv.cellprocessor.constraint.StrRegEx;
import org.supercsv.cellprocessor.ift.CellProcessor;
import org.supercsv.io.CsvBeanReader;
import org.supercsv.io.ICsvBeanReader;
import org.supercsv.prefs.CsvPreference;

/**
 * The InputHandler class receives the export from the parent Rapla application.  Using SuperCSV, ScheduleGenie
 * first parses the CSV input.  It sets up column headers based on the input, and places each shot into a 
 * TestShot object.  These objects are stored in a LinkedList.  The shots within the list are then separated 
 * by lab, and those shots are then separated by the day of the week.  The result is the basis for output from 
 * the ScheduleHandler class.
 * @author Kevin Tooley
 * @version 1.0.0
 */
public class InputHandler {
	
	// For testing purposes only; file will be read from Rapla export function
	//static final String CSV_FILENAME = "C:/Users/ktooley/Documents/ScheduleGenie_TEST/180822_Rev1.csv"; // TODO: Set filename to operator choice
	
	// Create Semicolon preference
	private static final CsvPreference SEMI_DELIMITED = new CsvPreference.Builder('"', ';', "\n").build();
	
	// LinkedList to hold test shots
	private LinkedList<TestShot> shotList = new LinkedList<TestShot>();  // Holds items parsed from CSV
	
	/**
	 * Using Super CSV, parse the input file (i.e. CSV export from Rapla).  TestShot members must match the inputs from the CSV.
	 * Additional fields will need to be handled with a code change.
	 * Additional resources are not an issue and will be part of the Resources string.
	 */
	public void parseCsv(String filename) {
		
		final String CSV_FILENAME = filename; 
		
		//try(ICsvBeanReader beanReader = new CsvBeanReader(new FileReader(CSV_FILENAME), CsvPreference.STANDARD_PREFERENCE))
		try(ICsvBeanReader beanReader = new CsvBeanReader(new FileReader(CSV_FILENAME), SEMI_DELIMITED))
        {
            // the header elements are used to map the values to the bean
            //final String[] headers = beanReader.getHeader(true);
			beanReader.getHeader(true);

			// Manually setting the header names because we don't want to read the last column
            final String[] headers = new String[]{"Name","Start","End","Resources","Persons", "duration", null};
            final CellProcessor[] processors = getProcessors();
 
            // Add test shots to the LinkedList
            TestShot testshot;
            while ((testshot = beanReader.read(TestShot.class, headers, processors)) != null) {
                shotList.add(testshot);
            }
        } catch (FileNotFoundException e) {
			// Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
     * Sets up the processors used for the examples.
     */
    private static CellProcessor[] getProcessors() {
        //final String emailRegex = "[a-z0-9\\._]+@[a-z0-9\\.]+";
        //StrRegEx.registerMessage(emailRegex, "must be a valid email address");
 
        final CellProcessor[] processors = new CellProcessor[] {
                new NotNull(), // Name
                new NotNull(), // Start
                new NotNull(), // End
                new NotNull(), // Resources
                new NotNull(), // Persons
                new NotNull(), // duration
                new Optional() // ExtraColumn
        };
        return processors;
    }

    /**
     * Return the list of shots contained in the csv file
     * @return LinkedList that contains the test shots
     */
	public LinkedList<TestShot> getShotList() {
		return shotList;
	}
	
	public void formatInput() {
		
	}

}