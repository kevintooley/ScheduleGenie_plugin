package org.rapla.plugin.schedulegenie;

import org.rapla.client.ClientServiceContainer;
import org.rapla.client.RaplaClientExtensionPoints;
import org.rapla.components.xmlbundle.I18nBundle;
import org.rapla.components.xmlbundle.impl.I18nBundleImpl;
import org.rapla.framework.Configuration;
import org.rapla.framework.PluginDescriptor;
import org.rapla.framework.RaplaContextException;
import org.rapla.framework.StartupEnvironment;
import org.rapla.framework.TypedComponentRole;
import org.rapla.plugin.export2ical.Export2iCalAdminOption;
import org.rapla.plugin.export2ical.Export2iCalMenu;
import org.rapla.plugin.export2ical.Export2iCalPlugin;
import org.rapla.plugin.export2ical.Export2iCalUserOption;
import org.rapla.plugin.export2ical.IcalPublicExtensionFactory;

public class TestScheduleGeniePlugin implements PluginDescriptor<ClientServiceContainer>{
	
	/*public static final TypedComponentRole<I18nBundle> RESOURCE_FILE = new TypedComponentRole<I18nBundle>(Export2iCalPlugin.class.getPackage().getName() + ".Export2iCalResources");
	
	public static final boolean ENABLE_BY_DEFAULT = false;
    public static final String EXPORT_ATTENDEES_PARTICIPATION_STATUS = "export_attendees_participation_status";

    public static final TypedComponentRole<String> EXPORT_ATTENDEES_PARTICIPATION_STATUS_PREFERENCE = new TypedComponentRole<String>("export_attendees_participation_status");

	public void provideServices(ClientServiceContainer container, Configuration config) throws RaplaContextException {
		container.addContainerProvidedComponent(RaplaClientExtensionPoints.PLUGIN_OPTION_PANEL_EXTENSION, Export2iCalAdminOption.class);
		if (!config.getAttributeAsBoolean("enabled", ENABLE_BY_DEFAULT))
			return;

		container.addResourceFile(RESOURCE_FILE);
	    container.addContainerProvidedComponent( RaplaClientExtensionPoints.PUBLISH_EXTENSION_OPTION, IcalPublicExtensionFactory.class);
	    
	    final int startupMode = container.getStartupEnvironment().getStartupMode();
        if ( startupMode != StartupEnvironment.APPLET)
        {
        	container.addContainerProvidedComponent(RaplaClientExtensionPoints.EXPORT_MENU_EXTENSION_POINT, Export2iCalMenu.class);
        }
	    container.addContainerProvidedComponent(RaplaClientExtensionPoints.USER_OPTION_PANEL_EXTENSION, Export2iCalUserOption.class);
	}*/
	
	public static final boolean ENABLE_BY_DEFAULT = false;
	public static final TypedComponentRole<I18nBundle> RESOURCE_FILE = new TypedComponentRole<I18nBundle>(TestScheduleGeniePlugin.class.getPackage().getName() + ".TestScheduleGeniePluginResources");

	@SuppressWarnings("deprecation")
	public void provideServices(ClientServiceContainer container, Configuration config) {
	  if ( !config.getAttributeAsBoolean("enabled", ENABLE_BY_DEFAULT) )
	     return;

	  container.addContainerProvidedComponent( RESOURCE_FILE, I18nBundleImpl.class,I18nBundleImpl.createConfig( RESOURCE_FILE.getId() ) );

		
		System.out.println("********** THE PLUGIN IS ENABLED ************");
		
	}

}
