package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import org.junit.Test;
import org.rapla.plugin.schedulegenie.InputHandler;

public class InputHandlerTest {

	@Test
	public void testParseCsv() {
		
		InputHandler ih = new InputHandler();
		ih.parseCsv();
		
		System.out.println("Shot List size = " + ih.getShotList().size());
		assertTrue(ih.getShotList().size() > 158);
		assertTrue(ih.getShotList().getFirst().getName().equals("MA - Overnight Automation"));
		assertTrue(ih.getShotList().getLast().getName().equals("OE Refresh"));
	}

}
