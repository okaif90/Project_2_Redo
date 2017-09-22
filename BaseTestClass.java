package com.amex.TS.AutomationTesting.GCO;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.testng.*;
import org.testng.annotations.*;

import com.saucelabs.common.SauceOnDemandAuthentication;
import com.saucelabs.common.SauceOnDemandSessionIdProvider;
import com.saucelabs.common.Utils;
import com.saucelabs.saucerest.SauceREST;
import com.saucelabs.testng.SauceOnDemandAuthenticationProvider;
import com.saucelabs.testng.SauceOnDemandTestListener;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.LocalFileDetector;
import org.openqa.selenium.remote.RemoteWebDriver;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.ThreadContext;

@Listeners({SauceOnDemandTestListener.class})
public abstract class BaseTestClass implements SauceOnDemandSessionIdProvider, SauceOnDemandAuthenticationProvider {

	private static final Logger logger = LogManager.getLogger();

	public String username = "default";
	public String accesskey = "default";
	private SauceREST sauceREST;

	public SauceOnDemandAuthentication authentication;

	private ThreadLocal<WebDriver> webDriver = new ThreadLocal<WebDriver>();
	private ThreadLocal<WebElement> object = new ThreadLocal<WebElement>();
	private ThreadLocal<String> sessionId = new ThreadLocal<String>();
	private ThreadLocal<String> excelTab = new ThreadLocal<String>();
	private ThreadLocal<String> excelPath = new ThreadLocal<String>();
	private ThreadLocal<String> testName = new ThreadLocal<String>();
	private ThreadLocal<String> currentUser = new ThreadLocal<String>();
	private ThreadLocal<String> buildTag = new ThreadLocal<String>();

	protected void setBuildTag(String newBuildTag){
		buildTag.set(newBuildTag);
	}

	public String getBuildTag(){
		return buildTag.get();
	}

	protected String setExcelPath(String newExcelFile){
		logger.debug("Excel file path being set to " + newExcelFile);
		excelPath.set(newExcelFile);
		return excelPath.get();
	}

	public String getTestName(){
		logger.trace("Getting test name");
		return testName.get();
	}

	protected String setTestName(String newTestName){
		logger.debug("Setting test name to " + newTestName);
		testName.set(newTestName);
		return testName.get();
	}

	public String getExcelPath(){
		logger.trace("Getting Excel path");
		return excelPath.get();
	}

	protected String setExcelTab(String tab){
		logger.debug("Setting Excel tab to " + tab);
		excelTab.set(tab);
		return getExcelTab();
	}

	protected WebElement getObject(){
		logger.trace("Getting the object");
		return object.get();
	}

	protected String setCurrentUser(String currentUserStr){
		logger.debug("Setting current user to " + currentUserStr);
		currentUser.set(currentUserStr);
		return currentUser.get();
	}

	public String getCurrentUser(){
		logger.trace("Getting current user");
		return currentUser.get();
	}

	@Override
	public SauceOnDemandAuthentication getAuthentication() {
		return authentication;
	}

	public WebDriver getWebDriver(){
		logger.trace("Getting the WebDriver");
		return webDriver.get();
	}

	public String getSessionId() {
		logger.trace("Getting the session ID");
		return sessionId.get();
	}

	public String getExcelTab(){
		logger.trace("Getting the Excel tab");
		return excelTab.get();
	}

	public WebElement setObject(WebElement newObject){
		logger.debug("Setting the object");
		object.set(newObject);
		return getObject();
	}

	protected WebDriver createDriver(String browser){
		logger.info("Creating a local " + browser + " WebDriver");
		switch(browser){
		case "internet explorer":
			logger.debug("Attempting to get the IEWebDriver");
			System.setProperty("webdriver.ie.driver",
					"C:/Users/" + System.getProperty("user.name") + "/Documents/MF Testing/Drivers/IEDriverServer.exe");
			webDriver.set(new InternetExplorerDriver());
			break;
		case "chrome":
			logger.debug("Attempting to get the ChromeWebDriver");
			System.setProperty("webdriver.chrome.driver",
					"C:/Users/" + System.getProperty("user.name") + "/Documents/MF Testing/Drivers/chromedriver.exe");
			logger.trace("Disabling Chrome extensions");
			ChromeOptions options = new ChromeOptions();
			options.addArguments("--disable-extensions");
			webDriver.set(new ChromeDriver(options));
			break;
		case "firefox":
			logger.debug("Creating a Firefox WebDriver");
			webDriver.set(new FirefoxDriver());
			break;
		default:
			logger.warn("Improper browser option given: " + browser);
			logger.info("Creating a Firefox WebDriver because no proper browser option was given");
			webDriver.set(new FirefoxDriver());
		}
		return webDriver.get();
	}

	protected WebDriver createDriver(String browser, String version, String os, String methodName, String tag) throws MalformedURLException{
		logger.info("Starting setup of remote " + browser + " WebDriver on a(n) " + os + "version " + version);
		DesiredCapabilities capabilities = new DesiredCapabilities();

		capabilities.setCapability(CapabilityType.BROWSER_NAME, browser);
		if(version != null){
			capabilities.setCapability(CapabilityType.VERSION, version);
		}
		capabilities.setCapability(CapabilityType.PLATFORM, os);
		logger.trace("Setting the maximum test duration to 3600 seconds");
		capabilities.setCapability("maxDuration", 3600);
		logger.trace("Setting the idle test timeout to 120 seconds");
		capabilities.setCapability("idleTimeout", 120);

		capabilities.setCapability("tags", tag);

		String jobName = methodName + '_' + os + '_' + browser + '_' + version;
		capabilities.setCapability("name", jobName);

		logger.info("Creating the remote WebDriver");
		RemoteWebDriver driver = new RemoteWebDriver(new URL("http://" + authentication.getUsername() + ":" + authentication.getAccessKey() + "@ondemand.saucelabs.com:80/wd/hub"), capabilities);
		driver.setFileDetector(new LocalFileDetector());
		webDriver.set(driver);

		String id = ((RemoteWebDriver)getWebDriver()).getSessionId().toString();
		sessionId.set(id);
		System.out.println(id);

		String message = String.format("SauceOnDemandSessionID=%1$s job-name=%2$s", id, jobName);
		logger.info(message);

		return webDriver.get();
	}

	@DataProvider(name="hardCodedBrowsers", parallel=true)
	public static Object[][] sauceBrowserDataProvider(Method testMethod){
		logger.debug("Starting dataprovider");
		Properties prop = null;
		prop = TMX_lib.propReader(prop);
		int numBrowsers = TMX_lib.getNumEles(prop.getProperty("Browser", "Firefox"));
		logger.debug("Number of browsers is " + numBrowsers);

		String[] browsers = prop.getProperty("Browser", "Firefox").split(",");
		String[] versions = prop.getProperty("Version", "40").split(",");

		Object[][] setup = new Object[numBrowsers][4];

		for(int i=0;i<numBrowsers;i++){
			setup[i] = new Object[]{browsers[i], versions[i], prop.getProperty("Platform", "Windows 7"), testMethod};
		}
		return setup;
	}

	@BeforeMethod(alwaysRun = true)
	public void setUp(Method testMethod) throws IOException{
		String testName = testMethod.getName();
		SauceOnDemandTestListener.verboseMode = false;
		ThreadContext.put("testName", testName);
		logger.debug("Setting up SauceLabs username and access key in case of remote test run");
		username = System.getenv("SAUCE_USER_NAME") !=null ? System.getenv("SAUCE_USER_NAME") : System.getenv("SAUCE_USERNAME");
		accesskey = System.getenv("SAUCE_API_KEY") !=null ? System.getenv("SAUCE_API_KEY") : System.getenv("SAUCE_ACCESS_KEY");
		authentication = new SauceOnDemandAuthentication(username, accesskey);
		setTestName(testName);
		this.sauceREST = new SauceREST(username, accesskey);
	}

	@AfterMethod(alwaysRun = true)
	public void tearDown(ITestResult result, Method testMethod) throws Exception {

		String testName = testMethod.getName();

		logger.debug("Test " + testName + " cleanup starting");

		SauceOnDemandSessionIdProvider sessionIdProvider = (SauceOnDemandSessionIdProvider) result.getInstance();
		String sessionId = sessionIdProvider.getSessionId();

		if(System.getProperty("remoteRun").equalsIgnoreCase("true")){
			try{
				Map<String, Object> updates = new HashMap<String, Object>();
				updates.put("passed", result.isSuccess());
				Utils.addBuildNumberToUpdate(updates);
				sauceREST.updateJobInfo(sessionId, updates);
			} catch(Exception e){
			}
		}

		try{
			File file = new File(getExcelPath());
			file.deleteOnExit();
		} catch(Exception e){
		}

		webDriver.get().quit();
		logger.debug("Test " + testName + " cleanup ending");

	} // end tearDown

}
