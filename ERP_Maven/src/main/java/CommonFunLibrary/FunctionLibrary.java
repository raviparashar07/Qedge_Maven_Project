package CommonFunLibrary;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileReader;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import Utilities.PropertyFileUtil;

public class FunctionLibrary {
	public static WebDriver driver;
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: startBrowser
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	public static WebDriver startBrowser(WebDriver driver) throws Throwable
	{
		if(PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("chrome"))
		{
			System.setProperty("webdriver.chrome.driver", "./CommonJars/chromedriver.exe");
			driver=new ChromeDriver();
			}
		else if(PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("firefox"))
		{
			
		}
		else if(PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("ie"))
		{
			
		}
		return driver;
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: openApplication
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	public static void openApplication(WebDriver driver) throws Throwable
	{
		driver.get(PropertyFileUtil.getValueForKey("Url"));
		driver.manage().window().maximize();
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: waitForElement
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	public static void waitForElement(WebDriver driver, String locatort, String locatorv, String mywait)
	{
		WebDriverWait wait=new WebDriverWait(driver, Integer.parseInt(mywait));
		if(locatort.equalsIgnoreCase("id"))
		{
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorv)));
		}
		else if(locatort.equalsIgnoreCase("name"))
		{
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.name(locatorv)));
		}
		else if(locatort.equalsIgnoreCase("xpath"))
		{
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorv)));
		}
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: typeAction
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	
	public static void typeAction(WebDriver driver, String locatort, String locatorv, String testdata)
	{
		if(locatort.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(locatorv)).clear();
			driver.findElement(By.id(locatorv)).sendKeys(testdata);
		}
		
		else if(locatort.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(locatorv)).clear();
			driver.findElement(By.name(locatorv)).sendKeys(testdata);
		}
		
		else if(locatort.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(locatorv)).clear();
			driver.findElement(By.xpath(locatorv)).sendKeys(testdata);
		}
		
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: clickAction
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	
	public static void clickAction(WebDriver driver, String locatort, String locatorv)
	{
		if(locatort.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(locatorv)).sendKeys(Keys.ENTER);
		}
		else if(locatort.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(locatorv)).click();
		}
		else if(locatort.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(locatorv)).click();
		}
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: closeBrowser
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	public static void closeBrowser(WebDriver driver)
	{
		driver.close();
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: generateDate
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */

public static String generateDate()
{
	Date date=new Date();
	SimpleDateFormat sdf=new SimpleDateFormat("YYYY_MM_dd_ss");
	return sdf.format(date);
}

/*Templates
 * ProjectName: ERP_Stock
 * Module Name: captureData
 * TesterName: Ravi Parashar
 * Creation Date:
 */
public static void captureData(WebDriver driver, String locatort, String locatorv) throws Throwable
{
	String snumber="";
	if(locatort.equalsIgnoreCase("id"))
	{
		snumber=driver.findElement(By.id(locatorv)).getAttribute("value");
	}
	else if(locatort.equalsIgnoreCase("xpath"))
	{
		snumber=driver.findElement(By.xpath(locatorv)).getAttribute("value");
	}
	else if(locatort.equalsIgnoreCase("name"))
	{
		snumber=driver.findElement(By.name(locatorv)).getAttribute("value");
	}
	FileWriter fw=new FileWriter("D:\\Ojt4oclock\\ERP_Maven\\CaptureData\\supplier.txt");
	BufferedWriter bw=new BufferedWriter(fw);
	bw.write(snumber);
	bw.flush();
	bw.close();
	
	
}

/*Templates
 * ProjectName: ERP_Stock
 * Module Name: tableValidation
 * TesterName: Ravi Parashar
 * Creation Date:
 */

public static void tableValidation(WebDriver driver, String columndata) throws Throwable
{
	FileReader fr=new FileReader("D:\\Ojt4oclock\\ERP_Maven\\CaptureData\\supplier.txt");
	BufferedReader br=new BufferedReader(fr);
	String Exp_data=br.readLine();
	
	//convert columndata into integer
	
	int column=Integer.parseInt(columndata);
	if(driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).isDisplayed())
	{
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).clear();
		Thread.sleep(3000);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).sendKeys(Exp_data);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button1"))).click();
		Thread.sleep(3000);
		
	}
	else 
	{    // click on search panel button
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-panel1"))).click();
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).clear();
		Thread.sleep(3000);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).sendKeys(Exp_data);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button1"))).click();
		Thread.sleep(3000);
	}
	
	// validation for table
	WebElement table=driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("sup-table")));
	List<WebElement>rows=table.findElements(By.tagName("tr"));
	
	for(int i=1; i<rows.size()-1; i++)
	{
		String act_data=driver.findElement(By.xpath("//table[@id='tbl_a_supplierslist']/tbody/tr["+i+"]/td["+column+"]/div/span/span")).getText();
		System.out.println("no of rows are::"+rows.size());
		Thread.sleep(5000);
		System.out.println(Exp_data+"  "+act_data);
		Assert.assertEquals(Exp_data, act_data,"snumber is not matching");
		break;
		
	}
}
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: StockCategories
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	
	public static void stockCategories(WebDriver driver) throws Throwable
	{
		Actions ac=new Actions(driver);
		WebElement stockitems=driver.findElement(By.xpath("//*[@id='mi_a_stock_items']/a"));
		ac.moveToElement(stockitems).perform();
		Thread.sleep(3000);
		WebElement stockc=driver.findElement(By.xpath("//*[@id='mi_a_stock_categories']/a"));
		ac.moveToElement(stockc).click().perform();
		Thread.sleep(3000);
		
	}
	
	/*Templates
	 * ProjectName: ERP_Stock
	 * Module Name: stockValidation
	 * TesterName: Ravi Parashar
	 * Creation Date:
	 */
	
	public static void stockValidation(WebDriver driver, String Exp_data) throws Throwable
	{
		if(driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).isDisplayed())
		{
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).clear();
			Thread.sleep(3000);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).sendKeys(Exp_data);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button1"))).click();
			Thread.sleep(3000);
			
		}
		else
		{
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-panel1"))).click();
			Thread.sleep(3000);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).clear();
			Thread.sleep(3000);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-box1"))).sendKeys(Exp_data);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button1"))).click();
			Thread.sleep(3000);
		}
		
		WebElement table=driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("sot-table")));
		List<WebElement>rows=table.findElements(By.tagName("tr"));
		System.out.println("no of rows are::"+rows.size());
		for(int i=1; i<=rows.size()-1; i++)
		{
			String Act_data=driver.findElement(By.xpath("//table[@id='tbl_a_stock_categorieslist']/tbody/tr["+i+"]/td[4]/div/span/span")).getText();
		System.out.println(Exp_data+"  "+Act_data);
		Assert.assertEquals(Exp_data, Act_data, "Data is not matching");
		break;
		}
	}
}






















































