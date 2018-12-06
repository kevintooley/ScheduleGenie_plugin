package org.rapla.plugin.schedulegenie;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;

public class Day {
	
	private final String date;
	private LinkedList<TestShot> shots = new LinkedList<TestShot>();
	
	public Day(String date) {
		super();
		this.date = date;
	}

	public LinkedList<TestShot> getShotList() {
		return shots;
	}

	public void addShots(TestShot shot) {
		shots.add(shot);
	}

	public Date getDate() throws ParseException {
		SimpleDateFormat ft = new SimpleDateFormat ("MM-dd-yyyy");
		
		return ft.parse(date);
	}
	
	
	

}
