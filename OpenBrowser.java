package Task13;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.chrome.ChromeDriver;

public class OpenBrowser {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		//setup driver and open browser
		ChromeDriver driver = new ChromeDriver();
		//Pass url to load in browser
		driver. get ("https://www.google.com/");
		//To maximize window of browser
		driver.manage().window().maximize();
		//To locate the search box using locator and enter text to be searched, and to select  enter we used keys tag
		driver.findElement(By.name("q")).sendKeys("Selenium browser driver",Keys.ENTER);
		//to close the browser
		
	}

}
