package org.rapla.plugin.schedulegenie;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedList;

import org.supercsv.cellprocessor.Optional;
//import org.supercsv.cellprocessor.ParseInt;
//import org.supercsv.cellprocessor.ParseLong;
import org.supercsv.cellprocessor.constraint.NotNull;
//import org.supercsv.cellprocessor.constraint.StrRegEx;
import org.supercsv.cellprocessor.ift.CellProcessor;
import org.supercsv.io.CsvBeanReader;
import org.supercsv.io.ICsvBeanReader;
import org.supercsv.prefs.CsvPreference;

/**
 * Using SuperCSV, ScheduleGenie parses the lab_configuration.cfg file to configure the string outputs
 * for the Bulk Upload spreadsheet
 * @author Kevin Tooley
 * @version 1.0.0
 */
public class InputHandler {
	
	// Create Semicolon preference
	private static final CsvPreference SEMI_DELIMITED = new CsvPreference.Builder('"', ';', "\n").build();
	
	// LinkedList to hold CPTS Lab to Test Site Scheduling System (TSSS) lab identifiers
	private LinkedList<LabMapping> LAB_MAPS = new LinkedList<LabMapping>();  // Holds items parsed from cfg	
	
	// LinkedList to hold the configuration mappings for the TSSS
	private LinkedList<ConfigMapping> CONFIG_MAPS = new LinkedList<ConfigMapping>();
	
	/**
	 * Parse the given config file using SuperCSV
	 * @param filename
	 */
	public void parseCsv(File filename) {
		
		final String CSV_FILENAME = filename.getAbsolutePath(); 
		
		//try(ICsvBeanReader beanReader = new CsvBeanReader(new FileReader(CSV_FILENAME), CsvPreference.STANDARD_PREFERENCE))
		try(ICsvBeanReader beanReader = new CsvBeanReader(new FileReader(CSV_FILENAME), SEMI_DELIMITED))
        {
            // the header elements are used to map the values to the bean
			beanReader.getHeader(true);
			
			if (filename.getName().contains("lab_configuration")) {
				// Manually setting the header names because we don't want to read the last column
	            final String[] headers = new String[]{"common_name","tsss_name1","tsss_name2","tsss_name3","tsss_name4","tsss_name5","tsss_name6","tsss_name7","tsss_name8","tsss_name9","tsss_name10"};
	            final CellProcessor[] processors = getProcessors();
	 
	            // Add test shots to the LinkedList
	            LabMapping lab;
	            while ((lab = beanReader.read(LabMapping.class, headers, processors)) != null) {
	            	LAB_MAPS.add(lab);
	            }
			}
			else if (filename.getName().contains("map")) {
				// Manually setting the header names because we don't want to read the last column
	            final String[] headers = new String[]{"common_config_name","tsss_config_name"};
	            final CellProcessor[] processors = getConfigProcessors();
	 
	            // Add test shots to the LinkedList
	            ConfigMapping config;
	            while ((config = beanReader.read(ConfigMapping.class, headers, processors)) != null) {
	            	CONFIG_MAPS.add(config);
	            }
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
     * Sets up the processors used for the lab_configuration
     */
    private static CellProcessor[] getProcessors() {
        final CellProcessor[] processors = new CellProcessor[] {
                new NotNull(), // COMMON_NAME
                new NotNull(), // TSSS_NAME1
                new Optional(), // TSSS_NAME2
                new Optional(), // TSSS_NAME3
                new Optional(), // TSSS_NAME4
                new Optional(), // TSSS_NAME5
                new Optional(), // TSSS_NAME6
                new Optional(), // TSSS_NAME7
                new Optional(), // TSSS_NAME8
                new Optional(), // TSSS_NAME9
                new Optional(), // TSSS_NAME10
        };
        return processors;
    }
    
    /**
     * Sets up the processors used for the configuration_mapping
     */
    private static CellProcessor[] getConfigProcessors() {
        final CellProcessor[] processors = new CellProcessor[] {
                new NotNull(), // common_config_name
                new NotNull(), // tsss_config_name
        };
        return processors;
    }

    /**
     * Return the list of lab to TSSS mappings contained in the cfg file
     * @return LinkedList that contains the mappings
     */
	public LinkedList<LabMapping> getMapping() {
		return LAB_MAPS;
	}
	
	/**
     * Return the list of lab to TSSS mappings contained in the cfg file
     * @return LinkedList that contains the mappings
     */
	public LinkedList<ConfigMapping> getConfigMapping() {
		return CONFIG_MAPS;
	}

}
