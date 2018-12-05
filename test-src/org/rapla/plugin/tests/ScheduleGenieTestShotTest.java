package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import org.rapla.plugin.schedulegenie.*;

import org.junit.Test;

public class ScheduleGenieTestShotTest {

	@Test
	public void testTestShot() {
		
		TestShot ts = new TestShot();
		
		assertNotNull(ts);
	}

	@Test
	public void testConstructor() {
		
		TestShot ts = new TestShot("string1", "string2", "string3", "string4", "string5", "string6");

		assertTrue(ts.getName().equals("string1"));
		assertTrue(ts.getStart().equals("string2"));
		assertTrue(ts.getEnd().equals("string3"));
		assertTrue(ts.getResources().equals("string4"));
		assertTrue(ts.getPersons().equals("string5"));
		assertTrue(ts.getDuration().equals("string6"));
		
	}

	@Test
	public void testSetName() {
		TestShot ts = new TestShot();
		ts.setName("string1");
		assertTrue(ts.getName().equals("string1"));
	}

	@Test
	public void testSetStart() {
		TestShot ts = new TestShot();
		ts.setStart("string1");
		assertTrue(ts.getStart().equals("string1"));
	}

	@Test
	public void testSetEnd() {
		TestShot ts = new TestShot();
		ts.setEnd("string1");
		assertTrue(ts.getEnd().equals("string1"));
	}

	@Test
	public void testSetResources() {
		TestShot ts = new TestShot();
		ts.setResources("string1");
		assertTrue(ts.getResources().equals("string1"));
	}

	@Test
	public void testSetPersons() {
		TestShot ts = new TestShot();
		ts.setPersons("string1");
		assertTrue(ts.getPersons().equals("string1"));
	}

	@Test
	public void testSetDuration() {
		TestShot ts = new TestShot();
		ts.setDuration("string1");
		assertTrue(ts.getDuration().equals("string1"));
	}

	@Test
	public void testToString() {
		TestShot ts = new TestShot("string1", "string2", "string3", "string4", "string5", "string6");
		assertTrue(ts.toString().equals("TestShot [ShotName=string1, Start=string2, End=string3, Resources=string4, Persons=string5, duration=string6]"));
	}

}
