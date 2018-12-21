package org.rapla.plugin.tests;

import static org.junit.Assert.*;

import java.io.File;
import org.junit.Test;
import org.rapla.plugin.schedulegenie.InputHandler;

public class InputHandlerTest {

	@Test
	public void testParseCsv() {
		
		System.out.println("Starting InputHandler test");
		
		InputHandler ih = new InputHandler();
		
		//ih.parseCsv(file);
		ih.parseCsv(new File(System.getProperty("user.dir") + System.getProperty("file.separator") + "data" + System.getProperty("file.separator") + "lab_configuration.cfg"));		        
		
		/*for (LabMapping	lab : ih.getMapping()) {
			System.out.println(lab.toString());
		}*/
		//System.out.println("Shot List size = " + ih.getMapping().size());
		
		assertTrue(ih.getMapping().size() == 6);
		assertTrue(ih.getMapping().get(0).getCommon_name().equals("AMOD1"));
		assertTrue(ih.getMapping().get(0).getTsss_name1().equals("AMOD NSCC TI12 SUITE 1 CND"));
		assertTrue(ih.getMapping().get(0).getTsss_name2().equals("AMOD NSCC TI12 SUITE 1 WCS"));
		assertTrue(ih.getMapping().get(0).getTsss_name3().equals("AMOD NSCC TI12 SUITE 1 SPY"));
		assertTrue(ih.getMapping().get(0).getTsss_name4().equals("AMOD NSCC TI12 SUITE 1 ADS"));
		assertTrue(ih.getMapping().get(0).getTsss_name5().equals("AMOD NSCC TI12 SUITE 1 ACTS"));
		assertTrue(ih.getMapping().get(0).getTsss_name6().equals("AMOD NSCC TI12 SUITE 1 ORTS"));
		assertTrue(ih.getMapping().get(0).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(0).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(0).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(0).getTsss_name10() == null);
		
		assertTrue(ih.getMapping().get(1).getCommon_name().equals("BL10_SUITE"));
		assertTrue(ih.getMapping().get(1).getTsss_name1().equals("NSCC BL10 CND"));
		assertTrue(ih.getMapping().get(1).getTsss_name2().equals("NSCC BL10 WCS"));
		assertTrue(ih.getMapping().get(1).getTsss_name3().equals("NSCC BL10 SPY"));
		assertTrue(ih.getMapping().get(1).getTsss_name4().equals("NSCC BL10 ADS"));
		assertTrue(ih.getMapping().get(1).getTsss_name5().equals("NSCC BL10 ACTS"));
		assertTrue(ih.getMapping().get(1).getTsss_name6().equals("NSCC BL10 ORTS"));
		assertTrue(ih.getMapping().get(1).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(1).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(1).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(1).getTsss_name10() == null);
		
		assertTrue(ih.getMapping().get(2).getCommon_name().equals("LBTS"));
		assertTrue(ih.getMapping().get(2).getTsss_name1().equals("LBTS BL10 CND"));
		assertTrue(ih.getMapping().get(2).getTsss_name2().equals("LBTS BL10 WCS"));
		assertTrue(ih.getMapping().get(2).getTsss_name3().equals("LBTS BL10 SPY"));
		assertTrue(ih.getMapping().get(2).getTsss_name4().equals("LBTS BL10 ADS"));
		assertTrue(ih.getMapping().get(2).getTsss_name5().equals("LBTS BL10 ACTS"));
		assertTrue(ih.getMapping().get(2).getTsss_name6().equals("LBTS BL10 ORTS"));
		assertTrue(ih.getMapping().get(2).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(2).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(2).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(2).getTsss_name10() == null);
		
		assertTrue(ih.getMapping().get(3).getCommon_name().equals("TI12H"));
		assertTrue(ih.getMapping().get(3).getTsss_name1().equals("NSCC TI12H CND"));
		assertTrue(ih.getMapping().get(3).getTsss_name2().equals("NSCC TI12H WCS"));
		assertTrue(ih.getMapping().get(3).getTsss_name3().equals("NSCC TI12H SPY"));
		assertTrue(ih.getMapping().get(3).getTsss_name4().equals("NSCC TI12H ADS"));
		assertTrue(ih.getMapping().get(3).getTsss_name5().equals("NSCC TI12H ACTS"));
		assertTrue(ih.getMapping().get(3).getTsss_name6().equals("NSCC TI12H ORTS"));
		assertTrue(ih.getMapping().get(3).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(3).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(3).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(3).getTsss_name10() == null);
		
		assertTrue(ih.getMapping().get(4).getCommon_name().equals("SUITE_B"));
		assertTrue(ih.getMapping().get(4).getTsss_name1().equals("SUITE B"));
		assertTrue(ih.getMapping().get(4).getTsss_name2() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name3() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name4() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name5() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name6() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(4).getTsss_name10() == null);
		
		assertTrue(ih.getMapping().get(5).getCommon_name().equals("TI16"));
		assertTrue(ih.getMapping().get(5).getTsss_name1().equals("NSCC TI16 CND"));
		assertTrue(ih.getMapping().get(5).getTsss_name2().equals("NSCC TI16 WCS"));
		assertTrue(ih.getMapping().get(5).getTsss_name3().equals("NSCC TI16 SPY"));
		assertTrue(ih.getMapping().get(5).getTsss_name4().equals("NSCC TI16 ADS"));
		assertTrue(ih.getMapping().get(5).getTsss_name5().equals("NSCC TI16 ACTS"));
		assertTrue(ih.getMapping().get(5).getTsss_name6().equals("NSCC TI16 ORTS"));
		assertTrue(ih.getMapping().get(5).getTsss_name7() == null);
		assertTrue(ih.getMapping().get(5).getTsss_name8() == null);
		assertTrue(ih.getMapping().get(5).getTsss_name9() == null);
		assertTrue(ih.getMapping().get(5).getTsss_name10() == null);
		
		System.out.println("Finished InputHandler test");
	}

}
