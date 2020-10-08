package org.rapla.plugin.metricsgenie;

import org.rapla.client.ClientServiceContainer;
import org.rapla.client.RaplaClientExtensionPoints;
import org.rapla.components.xmlbundle.I18nBundle;
import org.rapla.framework.Configuration;
import org.rapla.framework.PluginDescriptor;
import org.rapla.framework.RaplaContextException;
import org.rapla.framework.StartupEnvironment;
import org.rapla.framework.TypedComponentRole;
import org.rapla.plugin.metricsgenie.MetricsGenieMenu;
import org.rapla.plugin.metricsgenie.MetricsGeniePlugin;

public class MetricsGeniePlugin implements PluginDescriptor<ClientServiceContainer>{
	
	public static final TypedComponentRole<I18nBundle> RESOURCE_FILE = new TypedComponentRole<I18nBundle>(MetricsGeniePlugin.class.getPackage().getName() + ".MetricsGeniePluginResources");
	
	public static final boolean ENABLE_BY_DEFAULT = true;
    
	public void provideServices(ClientServiceContainer container, Configuration config) throws RaplaContextException {
		
		if (!config.getAttributeAsBoolean("enabled", ENABLE_BY_DEFAULT))
			return;

		container.addResourceFile(RESOURCE_FILE);
	    
	    final int startupMode = container.getStartupEnvironment().getStartupMode();
        if ( startupMode != StartupEnvironment.APPLET)
        {
        	container.addContainerProvidedComponent(RaplaClientExtensionPoints.EXPORT_MENU_EXTENSION_POINT, MetricsGenieMenu.class);
        }
	    
        System.out.println("********** THE PLUGIN IS ENABLED ************");
	}

}