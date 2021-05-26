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
	
	public static String writeResourceToFile(String resourceName, String certPath) throws Exception {
	    File outFile = new File(certPath + File.separator + resourceName);

	    if (outFile.exists() && outFile.isFile()) {
	        return outFile.getAbsolutePath();
	    }
	    InputStream resourceStream = TestUtil.class.getResourceAsStream(resourceName);
	    if (resourceStream == null) {
		    URLClassLoader urlClassLoader = new URLClassLoader(new URL[0], TestUtil.class.getClassLoader());
		    Thread.currentThread().setContextClassLoader(urlClassLoader);
		    URL url = urlClassLoader.findResource(resourceName);
		    if (url != null) {
		        URLConnection conn = url.openConnection();
		        if (conn != null) {
		            resourceStream = conn.getInputStream();
		        }
		    }
		    urlClassLoader.close();
	    }
	    if (resourceStream != null) {
	        Files.copy(resourceStream, outFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
	        return outFile.getAbsolutePath();
	    } else {
	        throw new Exception("Embedded Resource " + resourceName + " not found.");
	    }
	}

}
