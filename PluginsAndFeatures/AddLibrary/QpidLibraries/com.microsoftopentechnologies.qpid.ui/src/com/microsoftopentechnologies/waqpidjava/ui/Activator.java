package com.microsoftopentechnologies.waqpidjava.ui;

import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.ui.plugin.AbstractUIPlugin;
import org.osgi.framework.BundleContext;

/**
 * The activator class controls the plug-in life cycle.
 */
public class Activator extends AbstractUIPlugin {

	// The plug-in ID
	public static final String PLUGIN_ID = "com.microsoftopentechnologies.qpid.ui";

	// The shared instance
	private static Activator plugin;
	private static BundleContext bundleContext;
	public static boolean dplyState;

	/**
	 * The constructor.
	 */
	public Activator() {
	}

	/**
	 * Start method.
	 */
	public void start(BundleContext context) throws Exception {
		super.start(context);
		plugin = this;
	}

	/**
	 * Stop method.
	 */
	public void stop(BundleContext context) throws Exception {
		plugin = null;
		super.stop(context);
	}

	/**
	 * Returns the shared instance.
	 * @return the shared instance
	 */
	public static Activator getDefault() {
		return plugin;
	}

	/**
	 * Returns bundle context.
	 * @return BundleContext
	 */
	public static BundleContext getBundleContext() {
		return bundleContext;
	}

	/**
     * Logs a message and exception.
     *
     * @param message
     * @param excp : exception.
     */
    public void log(String message, Exception excp) {
        getLog().log(new Status(IStatus.ERROR, PLUGIN_ID, message, excp));
    }
    
    /**
     * Save the Deployment Assembly check box state
     * which will be used in case of immediate edit.
     * @param state
     */
    public static void setDeploymentAssemblyState(boolean state) {
    	dplyState = state;
    }
    
    /**
     * Returns Deployment Assembly check box state
     * which will be used in case of immediate edit.
     * @return
     */
    public static boolean geteploymentAssemblyState() {
    	return dplyState;
    }
}
