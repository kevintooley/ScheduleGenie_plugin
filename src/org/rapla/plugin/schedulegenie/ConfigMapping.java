package org.rapla.plugin.schedulegenie;

public class ConfigMapping {

	// The common_config_name will be equal to the name contained within the ScheduleGenie::Resources::Config objects
	private String common_config_name;
	
	// This maps to the "baseline" field in the bulk upload template
	private String tsss_config_name;
	
	/**
	 * Constructor
	 */
	public ConfigMapping() {}
	
	/**
	 * Constructor
	 * @param common_config_name
	 * @param tsss_config_name
	 */
	public ConfigMapping(String common_config_name, String tsss_config_name) {
		super();
		this.common_config_name = common_config_name;
		this.tsss_config_name = tsss_config_name;
	}

	public String getCommon_config_name() {
		return common_config_name;
	}

	public void setCommon_config_name(String common_config_name) {
		this.common_config_name = common_config_name;
	}

	public String getTsss_config_name() {
		return tsss_config_name;
	}

	public void setTsss_config_name(String tsss_config_name) {
		this.tsss_config_name = tsss_config_name;
	}

	@Override
	public String toString() {
		return "ConfigMapping [common_config_name=" + common_config_name + ", tsss_config_name=" + tsss_config_name
				+ "]";
	}
	
	
	
}
