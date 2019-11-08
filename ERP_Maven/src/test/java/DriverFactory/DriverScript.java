package DriverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.formula.ptg.FuncVarPtg;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunLibrary.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript {
	
	WebDriver driver;
	ExtentReports report;
	ExtentTest test;
	
	public void startTest() throws Throwable
	{
		//create reference object for accessing excelutil methods
		ExcelFileUtil excel=new ExcelFileUtil();
		
		//iterating all row in MasterCases Sheet
		for(int i=1; i<=excel.rowCount("MasterTestCases"); i++)
		{
			String ModuleStatus="";
			
			//write as executed for testcases which are flag Y
			
			if(excel.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y"))
			{
				
				//store module name into TCModule and TCModule holds all the test cases 
				String TCModule=excel.getCellData("MasterTestCases", i, 1);
				
				//generate report under your project
				report = new ExtentReports("./Reports/"+" "+TCModule+" "+FunctionLibrary.generateDate()+".html");
				
				
				//iterate all rows in TCModule
				for(int j=1; j<=excel.rowCount(TCModule); j++)
				{   
					//Test case executes here
					test=report.startTest(TCModule);      //it generate test report from here
					
					
					//read all columns in TCModule testcases
					String Description=excel.getCellData(TCModule, j, 0);
					String Object_Type=excel.getCellData(TCModule, j, 1);
					String Locator_Type=excel.getCellData(TCModule, j, 2);
					String Locator_Value=excel.getCellData(TCModule, j, 3);
					String Test_Data=excel.getCellData(TCModule, j, 4);
					
					try
					{
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver=FunctionLibrary.startBrowser(driver);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("openApplication"))
						{
							FunctionLibrary.openApplication(driver);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("closeBrowser"))
						{
								FunctionLibrary.closeBrowser(driver);
								test.log(LogStatus.INFO, Description);
						}
						
						else if(Object_Type.equalsIgnoreCase("captureData"))
						{
							FunctionLibrary.captureData(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
							
							
						}
						else if(Object_Type.equalsIgnoreCase("tableValidation"))
						{
							FunctionLibrary.tableValidation(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("stockCategories"))
							{ 
							FunctionLibrary.stockCategories(driver);
							test.log(LogStatus.INFO, Description);
							}
						else if(Object_Type.equalsIgnoreCase("stockValidation"))
						{
							FunctionLibrary.stockValidation(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
							
						}
						
						
						
						//write as pass into TCModule status column
						excel.setCellData(TCModule, j, 5, "Pass");
						test.log(LogStatus.PASS, Description);
						ModuleStatus="true";
						
					}
					catch(Exception e)
					{
						System.out.println(e.getMessage());
						
						//write as fail into TCModule  status column
						excel.setCellData(TCModule, j, 5, "Fail");
						test.log(LogStatus.FAIL, Description);
						
						//take screen shot and store into one variable
						File screen=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						
						//copy screen shot into local system
						FileUtils.copyFile(screen, new File("D:\\Ojt4oclock\\ERP_Accounting\\Screen\\"+Description+FunctionLibrary.generateDate()+".png"));
						
						ModuleStatus="false";
						break;
					}
					if(ModuleStatus.equalsIgnoreCase("True"))
					{
						excel.setCellData("MasterTestCases", i, 3, "Pass");
					}
					else if(ModuleStatus.equalsIgnoreCase("False"))
					{
						excel.setCellData("MasterTestCases", i, 3, "Fail");
					}
					report.flush();
					report.endTest(test);
					}
				    }
			
			else
			{
				//write as not executed for testcases which are flag N
				excel.setCellData("MasterTestCases", i, 3, "Not Executed");
			}
		}
	}

}



























