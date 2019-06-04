package utils;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

public class PropertiesUtils {

	public static String getPropertiesValue(String key) throws IOException {
		File file = new File("credentials.properties");
		FileReader fileReader = new FileReader(file);
		Properties properties = new Properties();
		properties.load(fileReader);
		String value = properties.getProperty(key);
		
		return value;
		
       }	
}
