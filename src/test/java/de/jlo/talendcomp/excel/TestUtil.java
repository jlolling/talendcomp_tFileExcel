package de.jlo.talendcomp.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLClassLoader;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class TestUtil {
	
	public static String writeResourceToFile(String resourceName, String certPath) throws IOException {
	    File outFile = new File(certPath + File.separator + resourceName);

	    if (outFile.isFile()) {
	        return outFile.getAbsolutePath();
	    }
	    InputStream resourceStream = null;
	    
	    // Java: In caso di JAR dentro il JAR applicativo 
	    URLClassLoader urlClassLoader = (URLClassLoader) TestUtil.class.getClassLoader();
	    URL url = urlClassLoader.findResource(resourceName);
	    if (url != null) {
	        URLConnection conn = url.openConnection();
	        if (conn != null) {
	            resourceStream = conn.getInputStream();
	        }
	    }
	    
	    if (resourceStream != null) {
	        Files.copy(resourceStream, outFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
	        return outFile.getAbsolutePath();
	    } else {
	        System.err.println("Embedded Resource " + resourceName + " not found.");
	    }
	    
	    return null;
	}   

}
