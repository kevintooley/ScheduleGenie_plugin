package org.rapla.plugin.tests;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

@RunWith(Suite.class)
@SuiteClasses({ DayTest.class, InputHandlerTest.class, ScheduleGenieTestShotTest.class, SpreadsheetHandlerTest.class })
public class AllTests {

}
