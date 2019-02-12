package com.mnt.webbanking.Test;

import com.mnt.webbanking.Data.*;
import com.mnt.webbanking.ReportUtils.ReportUtil;
import com.mnt.webbanking.Test.Keywords;
import com.mnt.webbanking.Utils.Resources;
import com.mnt.webbanking.Utils.TestUtils;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.events.EventFiringWebDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class TestController extends Resources{

	@BeforeClass
	public void initBrowser() throws IOException {
		Initialize();
	}
	
	
	

	@Test
	public void TestCaseController() throws NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, IOException {
		
		String startTime = TestUtils.now("dd.MMMM.yyyy hh.mm.ss aaa");
		ReportUtil.startTesting(System.getProperty("user.dir")+"//src//test//java//com//mnt//webbanking//Reports//index.html", startTime, "Test", "1.5");
		ReportUtil.startSuite("Suite1");
		String TCStatus="Pass";
		
		// loop through the test cases
		for(int TC=2;TC<=SuiteData.getRowCount("TestCases");TC++) {
			
			
			
			String TestCaseID = SuiteData.getCellData("TestCases", "TCID", TC);
			String RunMode = SuiteData.getCellData("TestCases", "RunMode", TC);
			String Reuse = SuiteData.getCellData("TestCases", "Reuse", TC);
			System.out.println("Reuuse "+Reuse);
			
			
			//System.out.println("2nnnnn "+ReuseVar1);
			
			if(RunMode.equals("Y")) {
				String TSStatus="Pass";
				//ChromeOptions options = new ChromeOptions();
			
				//dr= new ChromeDriver(options);
				
				
				//dr = new FirefoxDriver();
				
				
				DesiredCapabilities returnCapabilities = DesiredCapabilities.internetExplorer();
				returnCapabilities
						.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
				returnCapabilities.setCapability(InternetExplorerDriver.ENABLE_PERSISTENT_HOVERING, false);
				 String ieDriver = System.getProperty("user.dir") + "\\BrowserDrivers\\IEDriverServer32.exe";
				System.setProperty("webdriver.ie.driver", ieDriver);
				dr = new InternetExplorerDriver(returnCapabilities);
				
				
				driver = new EventFiringWebDriver(dr);
				driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			
				int rows = TestStepData.getRowCount(TestCaseID);
				if(rows<2) { 
					rows=2;
				}
				System.out.println("First "+TestCaseID);
				String ReuseTestCase="Reuse_";
				if(TestCaseID.startsWith(ReuseTestCase)) {
					String[] arrSplit = Reuse.split("\\|");
				
						
						 // loop through test data
							for(int TD=2;TD<=rows;TD++ ) {
							
								// loop through the test steps
								System.out.println("SuiteData.getRowCount(TestCaseID)"+SuiteData.getRowCount(TestCaseID));
								
								for(int TS=2;TS<=SuiteData.getRowCount(TestCaseID);TS++) {
									for (int i = 0; i < arrSplit.length; i++) {
										 //System.out.println(arrSplit[i]);
										 String ReuseVar2 = arrSplit[i];
									
									keyword = SuiteData.getCellData(TestCaseID, "Keyword", TS);
									webElement = SuiteData.getCellData(TestCaseID, "WebElement", TS);
									ProceedOnFail = SuiteData.getCellData(TestCaseID, "ProceedOnFail", TS);
									TSID = SuiteData.getCellData(TestCaseID, "TSID", TS);
									Description = SuiteData.getCellData(TestCaseID, "Description", TS);
									TestDataField = SuiteData.getCellData(TestCaseID, "TestDataField", TS);
									String s1=TestDataField;
								if (ReuseVar2.contains(s1)) {
									// TestDataField = SuiteData.getCellData(TestCaseID, "TestDataField", TS);
									TestData = TestStepData.getCellData(TestCaseID, ReuseVar2, TD);
								Method method = Keywords.class.getMethod(keyword);
								TSStatus = (String) method.invoke(method);
								if (TSStatus.contains("Failed")) {
									// take the screenshot
									String filename = TestCaseID + "[" + (TD - 1) + "]" + TSID + "[" + TestData + "]";
									TestUtils.getScreenShot(filename);

									TCStatus = TSStatus;
									// report error
									ReportUtil.addKeyword(Description, keyword, TSStatus,
											"Screenshot/" + filename + ".jpg");

									if (ProceedOnFail.equals("N")) {
										break;
									}
								} else {
									ReportUtil.addKeyword(Description, keyword, TSStatus, null);
								}
								ReportUtil.addTestCase(TestCaseID, startTime,
										TestUtils.now("dd.MMMM.yyyy hh.mm.ss aaa"), TCStatus);

							}
									}

						}
										


					}
				}
				else {
				// loop through test data
				for(int TD=2;TD<=rows;TD++ ) {
				
					// loop through the test steps
					System.out.println("SuiteData.getRowCount(TestCaseID)"+SuiteData.getRowCount(TestCaseID));
					
					for(int TS=2;TS<=SuiteData.getRowCount(TestCaseID);TS++) {
						
						keyword = SuiteData.getCellData(TestCaseID, "Keyword", TS);
						webElement = SuiteData.getCellData(TestCaseID, "WebElement", TS);
						ProceedOnFail = SuiteData.getCellData(TestCaseID, "ProceedOnFail", TS);
						TSID = SuiteData.getCellData(TestCaseID, "TSID", TS);
						Description = SuiteData.getCellData(TestCaseID, "Description", TS);
						TestDataField = SuiteData.getCellData(TestCaseID, "TestDataField", TS);
								TestData = TestStepData.getCellData(TestCaseID, TestDataField, TD);
								Method method = Keywords.class.getMethod(keyword);	
								TSStatus = (String) method.invoke(method);
								
								if(TSStatus.contains("Failed")) {
									// take the screenshot
									String filename = TestCaseID+"["+(TD-1)+"]"+TSID+"["+TestData+"]";
									TestUtils.getScreenShot(filename);
									
									TCStatus=TSStatus;
									// report error
									ReportUtil.addKeyword(Description, keyword, TSStatus, "Screenshot/"+filename+".jpg");

									if(ProceedOnFail.equals("N")) {
										break;
									}
								} else {
									ReportUtil.addKeyword(Description, keyword, TSStatus, null);
								}
						
					ReportUtil.addTestCase(TestCaseID, startTime, TestUtils.now("dd.MMMM.yyyy hh.mm.ss aaa"), TCStatus);
					//driver.quit();
				}
			}
				}
			}
		}

		ReportUtil.endSuite();
		ReportUtil.updateEndTime(TestUtils.now("dd.MMMM.yyyy hh.mm.ss aaa"));
		
	}
	
	
	@AfterClass
	public void quitBrowser() {
		System.out.println("In quitBrowser---------------------------");
		//driver.quit();
	}
	
	
}
