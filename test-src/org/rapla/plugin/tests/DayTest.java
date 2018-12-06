package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.junit.Test;
import org.rapla.plugin.schedulegenie.Day;
import org.rapla.plugin.schedulegenie.TestShot;

public class DayTest {

	@Test
	public void testDay() {
		Day day = new Day("10-31-2018");
		
		SimpleDateFormat ft = new SimpleDateFormat ("MM-dd-yyyy");
		String dateStr = "10-31-2018";
		
		Date date = null;
		try {
			date = ft.parse(dateStr);
		} catch (ParseException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
		try {
			assertTrue(day.getDate().compareTo(date) == 0);
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	@Test
	public void testGetShotList() {
		TestShot shot1 = new TestShot("string1", "string2", "string3", "string4", "string5", "string6");		
		TestShot shot2 = new TestShot("string7", "string8", "string9", "string10", "string11", "string12");
		TestShot shot3 = new TestShot("string13", "string14", "string15", "string16", "string17", "string18");
		
		Day today = new Day("12/5/2018");
		today.addShots(shot1);
		today.addShots(shot2);
		today.addShots(shot3);
		
		assertTrue(today.getShotList().size() == 3);
		assertTrue(today.getShotList().get(1).getName().equals("string7"));

	}


}
