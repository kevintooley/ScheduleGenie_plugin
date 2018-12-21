package org.rapla.plugin.schedulegenie;

public class LabMapping {
	
	// The labName will be equal to the name contained within the ScheduleGenie::Resources::Suite objects
	private String common_name;
	
	// The field<number> strings will contain all the entries to be entered into the bulk upload spreadsheet
	// i.e. "AMOD NSCC TI12 SUITE 1 CND", "AMOD NSCC TI12 SUITE 1 WCS", etc.
	// Note:  The TSSS API is very inconsistent, and not all labs use the above convention.
	// Therefore, field population will vary
	private String tsss_name1;
	private String tsss_name2;
	private String tsss_name3;
	private String tsss_name4;
	private String tsss_name5;
	private String tsss_name6;
	private String tsss_name7;
	private String tsss_name8;
	private String tsss_name9;
	private String tsss_name10;
	
	/**
	 * Constructor
	 */
	public LabMapping() {}
	
	/**
	 * Constructor
	 * @param common_name
	 * @param tsss_name1
	 * @param tsss_name2
	 * @param tsss_name3
	 * @param tsss_name4
	 * @param tsss_name5
	 * @param tsss_name6
	 * @param tsss_name7
	 * @param tsss_name8
	 * @param tsss_name9
	 * @param tsss_name10
	 */
	public LabMapping(String common_name, String tsss_name1, String tsss_name2, String tsss_name3, String tsss_name4,
			String tsss_name5, String tsss_name6, String tsss_name7, String tsss_name8, String tsss_name9,
			String tsss_name10) {
		super();
		this.common_name = common_name;
		this.tsss_name1 = tsss_name1;
		this.tsss_name2 = tsss_name2;
		this.tsss_name3 = tsss_name3;
		this.tsss_name4 = tsss_name4;
		this.tsss_name5 = tsss_name5;
		this.tsss_name6 = tsss_name6;
		this.tsss_name7 = tsss_name7;
		this.tsss_name8 = tsss_name8;
		this.tsss_name9 = tsss_name9;
		this.tsss_name10 = tsss_name10;
	}

	public String getCommon_name() {
		return common_name;
	}

	public void setCommon_name(String common_name) {
		this.common_name = common_name;
	}

	public String getTsss_name1() {
		return tsss_name1;
	}

	public void setTsss_name1(String tsss_name1) {
		this.tsss_name1 = tsss_name1;
	}

	public String getTsss_name2() {
		return tsss_name2;
	}

	public void setTsss_name2(String tsss_name2) {
		this.tsss_name2 = tsss_name2;
	}

	public String getTsss_name3() {
		return tsss_name3;
	}

	public void setTsss_name3(String tsss_name3) {
		this.tsss_name3 = tsss_name3;
	}

	public String getTsss_name4() {
		return tsss_name4;
	}

	public void setTsss_name4(String tsss_name4) {
		this.tsss_name4 = tsss_name4;
	}

	public String getTsss_name5() {
		return tsss_name5;
	}

	public void setTsss_name5(String tsss_name5) {
		this.tsss_name5 = tsss_name5;
	}

	public String getTsss_name6() {
		return tsss_name6;
	}

	public void setTsss_name6(String tsss_name6) {
		this.tsss_name6 = tsss_name6;
	}

	public String getTsss_name7() {
		return tsss_name7;
	}

	public void setTsss_name7(String tsss_name7) {
		this.tsss_name7 = tsss_name7;
	}

	public String getTsss_name8() {
		return tsss_name8;
	}

	public void setTsss_name8(String tsss_name8) {
		this.tsss_name8 = tsss_name8;
	}

	public String getTsss_name9() {
		return tsss_name9;
	}

	public void setTsss_name9(String tsss_name9) {
		this.tsss_name9 = tsss_name9;
	}

	public String getTsss_name10() {
		return tsss_name10;
	}

	public void setTsss_name10(String tsss_name10) {
		this.tsss_name10 = tsss_name10;
	}

	@Override
	public String toString() {
		return "LabMapping [common_name=" + common_name + ", tsss_name1=" + tsss_name1 + ", tsss_name2=" + tsss_name2
				+ ", tsss_name3=" + tsss_name3 + ", tsss_name4=" + tsss_name4 + ", tsss_name5=" + tsss_name5
				+ ", tsss_name6=" + tsss_name6 + ", tsss_name7=" + tsss_name7 + ", tsss_name8=" + tsss_name8
				+ ", tsss_name9=" + tsss_name9 + ", tsss_name10=" + tsss_name10 + "]";
	}
	
	
	
}
