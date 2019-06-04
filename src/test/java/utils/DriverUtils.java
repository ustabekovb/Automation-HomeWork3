package utils;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.opera.OperaDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DriverUtils {
	
static WebDriver driver;
	
	public static WebDriver getDriver(String browser) {
		
		if(driver!=null) {             //singleton pattern
			return driver;
		}
		
		switch(browser.toLowerCase())){
		
			case "chrome":
				WebDriverManager.chromedriver().setup();
				driver = new ChromeDriver();
				break;
			case "firefox":
				WebDriverManager.firefoxdriver().setup();
				driver = new FirefoxDriver();
				break;
			case "opera":
				WebDriverManager.operadriver().setup();
				driver = new OperaDriver();
				break;
			case "ie":
				WebDriverManager.iedriver().setup();
				driver = new InternetExplorerDriver();
				break;
			default:
				WebDriverManager.chromedriver().setup();
				driver = new ChromeDriver();
				break;
		}
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		
		return driver;
		
	}

}
