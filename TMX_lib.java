/* 
 * This java file contains methods required to run the automation scripts. Place it in your TMX scripts folder in your
 * Eclipse workspace.
 * There are custom methods written to derive data from excel and perform the required actions. 
 * Add new methods to this file when you require another custom action. In TMX when you create the corresponding action
 * only include a call to the method written in this file and pass the label (matching with excel datasheet) as an argument
 * The action method in turn should call the readExcelStatic() method to get the data. 
 * When a new method is written and the action is updated in TMX, update the Action template Agile Designer file with the 
 * new action block.
 */

package com.amex.TS.AutomationTesting.GCO;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.SystemUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.LocalFileDetector;
import org.openqa.selenium.remote.RemoteWebElement;
import org.openqa.selenium.security.UserAndPassword;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class TMX_lib {

	static Calendar cal;
	static Date time;
	static int timeDiff = 0;
	static String dateFormat = "M/d/yyyy";
	static SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);

	private static final Logger logger = LogManager.getLogger("TMX_lib");

	public static String xlname;

	public static ArrayList<String> strArrayBuffer;

	/**
	 * Sets the path to the Excel data sheet.
	 * 
	 */
	public static String setxlpath(String browser, String testName) {
		// If not running on mac
		logger.info("Excel data file to be copied is " + xlname);
		logger.info("Copying the Excel data file and naming the copy to " + xlname + "_" + testName);
		String fileOutLoc = null;

		try{
			if (!SystemUtils.IS_OS_MAC) {
				logger.trace("Using the Windows/Unix-based filepath");
				ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
				fileOutLoc = System.getProperty("user.dir") + File.separator
						+ xlname + "_" + testName + "_" + browser + ".xlsx";
				logger.trace("fileOutLoc is " + fileOutLoc);
				File fileOut = new File(fileOutLoc);
				FileOutputStream outputStream = new FileOutputStream(fileOut);
				InputStream is = classLoader.getResourceAsStream(xlname + ".xlsx");
				logger.trace("Creating a workboook and writing to " + fileOutLoc + " to initialize the file");
				XSSFWorkbook databook = new XSSFWorkbook(is);
				databook.write(outputStream);
				outputStream.close();
				databook.close();
				logger.trace("Closed the Excel file");
			}
			// if running on Mac
			else {
				logger.trace("Using the Mac-based filepath");
				ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
				InputStream is = classLoader.getResourceAsStream(xlname + ".xlsx");
				fileOutLoc = System.getProperty("user.dir") + File.separator
						+ xlname + "_" + testName + "_" + browser + ".xlsx";
				logger.trace("fileOutLoc is " + fileOutLoc);
				File fileOut = new File(fileOutLoc);
				logger.trace("Creating a workboook and writing to " + fileOutLoc + " to initialize the file");
				XSSFWorkbook databook = new XSSFWorkbook(is);

				FileOutputStream outputStream = new FileOutputStream(fileOut);
				databook.write(outputStream);
				outputStream.close();
				databook.close();
				logger.trace("Closed the Excel file");
			}
		} catch(IOException e){
			logger.error("IOException occurred when creating the temporary data sheet to work from");
			Assert.fail("IOException occurred");
		}

		return fileOutLoc;
	}

	/**
	 * Takes in an object ID and checks whether that object exists. This method
	 * is able to check for absolute id, name, xpath, label, and image. This
	 * method is also able to check for partial linktext, href, id, name, value,
	 * and alt.
	 * 
	 * @param objectid
	 *            the id of an object to search for.
	 * @return returns true if the element is found, false if it is not found.
	 */
	public static WebElement allCheck(String objectid, String identifier, WebDriver driver) {
		logger.debug("Attempting to locate " + objectid + " with identifier " + identifier);
		boolean element = false;
		WebElement object = null;
		switch (identifier) {
		case "full class":
			object = elementExists(By.className(objectid), driver);
			break;
		case "full id":
			object = elementExists(By.id(objectid), driver);
			break;
		case "full name":
			object = elementExists(By.name(objectid), driver);
			break;
		case "full label":
			object = elementExists(By.xpath("id(//label[text() = '" + objectid + "']/@for)"), driver);
			break;
		case "full xpath":
			object = elementExists(By.xpath(objectid), driver);
			break;
		case "full linktext":
			object = elementExists(By.linkText(objectid), driver);
			break;
		case "full value":
			object = elementExists(By.xpath("//*[@value='" + objectid + "']"), driver);
			break;
		case "full title":
			object = elementExists(By.xpath("//*[@title='" + objectid + "']"), driver);
			break;
		case "partial label":
			object = elementExists(By.xpath("id(//label[contains(text(), '" + objectid + "')]/@for)"), driver);
			break;
		case "partial value":
			object = elementExists(By.xpath("//*[contains(@value, \"" + objectid + "\")]"), driver);
			break;
		case "partial name":
			object = elementExists(By.xpath("//*[contains(@name, \"" + objectid + "\")]"), driver);
			break;
		case "partial id":
			object = elementExists(By.xpath("//*[contains(@id, \"" + objectid + "\")]"), driver);
			break;
		case "partial linktext":
			object = elementExists(By.partialLinkText(objectid), driver);
			break;
		case "partial href":
			object = elementExists(By.xpath("//*[contains(@href, \"" + objectid + "\")]"), driver);
			break;
		case "full image":
			object = elementExists(By.xpath("//img[@src='/resource/" + objectid + "']"), driver);
			break;
		case "full alt":
			object = elementExists(By.xpath("//*[@alt='" + objectid + "']"), driver);
		case "partial alt":
			object = elementExists(By.xpath("//*[contains(@alt, \"" + objectid + "\")]"), driver);
			break;
		case "partial title":
			object = elementExists(By.xpath("//*[contains(@title, \"" + objectid + "\")]"), driver);
			break;
		case "none":
			logger.warn("No element identifier given");
		default:
			if(!identifier.equalsIgnoreCase("none")){
				logger.warn("Improper element identifier used- " + identifier);
			}
			if(element == false){
				object = elementExists(By.id(objectid), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full ID");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[@alt='" + objectid + "']"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full alt");
				}
			}
			if(element == false){
				object = elementExists(By.name(objectid), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full name");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("id(//label[text() = '"+objectid+"']/@for)"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full label");
				}
			}
			if(element == false){
				object = elementExists(By.xpath(objectid), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full xpath");
				}
			}
			if(element == false){
				object = elementExists(By.linkText(objectid), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full linktext");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[@value='" + objectid + "']"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full value");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[@title='" + objectid + "']"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full title");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@value, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial value");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("id(//label[contains(text(), '" + objectid + "')]/@for)"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial label");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@name, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial name");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@id, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial id");
				}
			}
			if(element == false){
				object = elementExists(By.partialLinkText(objectid), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial linktext");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@href, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial href");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//img[@src='/resource/"+objectid+"']"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using full src");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@alt, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial alt");
				}
			}
			if(element == false){
				object = elementExists(By.xpath("//*[contains(@title, \"" + objectid + "\")]"), driver);
				element = object != null;
				if(element == true){
					logger.info("Object " + objectid + " found by using partial title");
				}
			}
			if(element == false){
				logger.warn("Element " + objectid + " was not found.");
			}
		}
		return object;
	}

	public static WebElement allCheck(String objectid, WebDriver driver) {
		logger.debug("Attempting to locate " + objectid + " without an identifier");
		boolean element = false;
		WebElement object = null;

		if(element == false){
			object = elementExists(By.id(objectid), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full id");
			}
		}
		if(element == false){
			object = elementExists(By.name(objectid), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full name");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("id(//label[text() = '"+objectid+"']/@for)"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial id");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[@alt='" + objectid + "']"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full alt");
			}
		}
		if(element == false){
			object = elementExists(By.xpath(objectid), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full xpath");
			}
		}
		if(element == false){
			object = elementExists(By.linkText(objectid), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full linktext");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[@value='" + objectid + "']"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full value");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[@title='" + objectid + "']"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full title");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@value, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial value");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@name, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial name");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@id, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial id");
			}
		}
		if(element == false){
			object = elementExists(By.partialLinkText(objectid), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial linktext");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@href, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial href");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//img[@src='/resource/"+objectid+"']"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using full img");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@alt, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial alt");
			}
		}
		if(element == false){
			object = elementExists(By.xpath("//*[contains(@title, \"" + objectid + "\")]"), driver);
			element = object != null;
			if(element == true){
				logger.info("Element " + objectid + " found using partial title");
			}
		}
		if(element == false){
			logger.warn("Element " + objectid + " was not found");
		}

		return object;
	}

	private static WebElement locateElementSwitch(String idType, String identifier, WebDriver driver) {

		logger.debug("Attempting to locate " + identifier + " with identifier " + idType);
		WebElement object = null;

		switch (idType) {
		case "xpath":
			object = driver.findElement(By.xpath(identifier));
			break;
		case "ID":
			object = driver.findElement(By.id(identifier));
			break;
		case "name":
			object = driver.findElement(By.name(identifier));
			break;
		case "linkText":
			object = driver.findElement(By.linkText(identifier));
			break;
		case "title":
			object = driver.findElement(By.xpath("//*[@title='" + identifier + "']"));
			break;
		case "value":
			object = driver.findElement(By.xpath("//*[@value='" + identifier + "']"));
			break;
		case "partialLinkText":
			object = driver.findElement(By.partialLinkText(identifier));
			break;
		case "partialValue":
			object = driver.findElement(By.xpath("//*[contains(@value, \"" + identifier + "\")]"));
			break;
		case "partialName":
			object = driver.findElement(By.xpath("//*[contains(@name, \"" + identifier + "\")]"));
			break;
		case "partialID":
			object = driver.findElement(By.xpath("//*[contains(@id, \"" + identifier + "\")]"));
			break;
		case "partialHref":
			object = driver.findElement(By.xpath("//*[contains(@href, \"" + identifier + "\")]"));
			break;
		case "partialAlt":
			object = driver.findElement(By.xpath("//*[contains(@alt, \"" + identifier + "\")]"));
			break;
		case "partialTitle":
			object = driver.findElement(By.xpath("//*[contains(@title, \"" + identifier + "\")]"));
			break;
		case "label":
			object = driver.findElement(By.xpath("id(//label[text() = '" + identifier + "']/@for)"));
			break;
		case "image":
			object = driver.findElement(By.xpath("//img[@src='/resource/" + identifier + "']"));
			break;
		default:
			logger.error("Attempted to use invalid identifier type " + idType);
			Assert.fail("Attempted to use an invalid identifier type");
		}

		return object;
	}

	public static Properties propReader(Properties prop){
		logger.trace("About to read the Properties file");
		InputStream input = null;
		prop = new Properties();

		try{
			String filename = "config.properties";
			logger.trace("Trying to read the properties file");
			input = TMX_lib.class.getClassLoader().getResourceAsStream(filename);
			logger.trace("Properties file found.  Loading file into Propterties object");
			prop.load(input);
			logger.trace("Properties file loaded");
			input.close();
		} catch (FileNotFoundException e) {
			System.out.println("Oops");
			logger.error("FileNotFound exception occurred");
			logger.error(e.getStackTrace());
		} catch (IOException e) {
			System.out.println("Whoops");
			logger.error("IOException occurred");
			logger.error(e.getStackTrace());
		}
		return prop;
	}

	public static int getNumEles(String eles){
		String[] eleArray = eles.split(",");
		logger.debug("Array length is " + eleArray.length);
		return eleArray.length;
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * function verifies that a date matches the value used from the Excel
	 * sheet.
	 * 
	 * @param field_Label
	 */
	public static void verifyDateTime(String field_Label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying date and/or time");
		String sArg = null;
		sArg = readExcel(field_Label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		String objectid = strArray[1];
		logger.trace("Object id is " + objectid);
		String identifier = strArray[2];
		logger.trace("Object identifier is " + identifier);

		String actual = null;

		boolean isFound = false;
		logger.trace("Searching for object " + objectid);
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			logger.trace("Object " + objectid + " found");
			isFound = true;
			actual = object.getText();
			logger.debug("Object " + objectid + " text is '" + actual + "'");
		}

		if (!isFound) {
			logger.error("Object " + objectid + " not found!");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.fail("Element " + objectid + " not found.");
		}

		String timeDateToCheck;
		logger.trace("Pulling expected time from datasheet");
		if (strArrayBuffer.size() >= 1) {
			timeDiff = Integer.parseInt(strArrayBuffer.get(1));
			timeDateToCheck = strArrayBuffer.get(0);
		} else {
			timeDateToCheck = strArray[0];
		}

		Calendar adjuster = Calendar.getInstance();
		try {
			adjuster.setTime(sdf.parse(timeDateToCheck));
		} catch (ParseException e) {
			e.printStackTrace();
		}

		adjuster.add(Calendar.HOUR, timeDiff);
		timeDateToCheck = sdf.format(adjuster.getTime());

		if (!actual.contains(timeDateToCheck)) {
			logger.error("Object did not contain the correct time!");
			logger.error("Expected time: " + timeDateToCheck);
			logger.error("  Actual time: " + actual);
		}
		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method verifies whether an element on the screen contains the text pulled
	 * from the Excel sheet.
	 * 
	 * @param field_label
	 *            A string representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void verifyText(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying " + field_label + " on the page");
		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		logger.debug("Searching for object " + objectid);
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			logger.trace("Object " + objectid + " found");
			isFound = true;
			String actual = object.getText();
			logger.debug("Verifying text of object " + objectid);
			if (!testdata.contains("strArrayBuffer")) {
				if (!actual.contains(testdata)) {
					logger.error("Object " + objectid + " had incorrect text!");
					logger.error("Expected text: " + testdata);
					logger.error("  Actual text: " + actual);
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertEquals(actual, testdata, "Text expected: '" + testdata + "', actual: '" + actual + "'");
				}
				logger.debug("Expected text: " + testdata);
				logger.debug("  Actual text: " + actual);
			} else {
				//TODO add strArrayBuffer logging
				boolean isValidated = false;
				String expectedStr = null;
				for (String str : strArrayBuffer) {
					if (expectedStr == null)
						expectedStr = str;
					else
						expectedStr = expectedStr + ", " + str;
					if (actual.equalsIgnoreCase(str)) {
						isValidated = true;
						break;
					}
				}
				if (!isValidated) {
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertEquals(actual, expectedStr, "Expected one of '" + expectedStr + "', actual '" + actual + "'");
				} else {
				}
			}
		}
		if (!isFound) {
			logger.error("object " + objectid + " not found");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.fail("Object " + objectid + " not found, could not be verified.");
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	public static void verifyTextFromAD(String identifier, String testName, String step, String idType, String textToVerify, WebDriver driver) {

		logger.info("Verifying " + identifier + " on the page from TCO");
		WebElement object = locateElementSwitch(idType, identifier, driver);

		String actual = object.getText();
		if (!actual.contains(textToVerify)) {
			logger.error("Object " + identifier + " had incorrect text!");
			logger.error("Expected text: " + textToVerify);
			logger.error("  Actual text: " + actual);
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.fail("Text expected: '" + textToVerify + "', actual: '" + actual + "'");
		}

		logger.info("Expected text: " + textToVerify);
		logger.info("  Actual text: " + actual);

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	public static void verifyImg(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying picture " + field_label + " on the page");
		logger.info("This does not verify the contents of the image, just that the image is present!");
		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			String actual = object.getAttribute("src");
			if (!testdata.contains("strArrayBuffer")) {
				if (!actual.contains(testdata)) {
					logger.error("Picture source incorrect!");
					logger.error("Expected source: " + testdata);
					logger.error("  Actual source: " + actual);
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertEquals(actual, testdata,
							"Text expected: '" + testdata + "', actual :'" + actual + "'");
				}
				logger.info("Expected source: " + testdata);
				logger.info("  Actual source: " + actual);
			} else {
				//TODO implement strArrayBuffer logging
				boolean isValidated = false;
				String expectedStr = null;
				for (String str : strArrayBuffer) {
					if (expectedStr == null)
						expectedStr = str;
					else
						expectedStr = expectedStr + ", " + str;
					if (actual.equalsIgnoreCase(str)) {
						isValidated = true;
						break;
					}
				}
				if (!isValidated) {
					//status = "Fail";
					//comment = "Expected one of '" + expectedStr + "', actual '" + actual + "'";
					// LogStep();
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertEquals(actual, expectedStr,
							"Expected one of '" + expectedStr + "', actual '" + actual + "'");
				} else {
					//comment = "Element text found is '" + actual + "'";
				}
			}
		}

		if (!isFound) {
			logger.error("Object " + objectid + " not found");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.assertEquals(true, false, "Object not found, could not be verified.");
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	/**
	 * This method takes in a String representing a label on the Excel data
	 * sheet. This method verifies that the text pulled from the data sheet is
	 * not present on the page.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet
	 * @param object 
	 */
	public static void verifyTextNotPresent(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying " + field_label + " is not present on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			String actual = object.getText();
			System.out.println(actual);
			if (!testdata.contains("strArrayBuffer")) {
				if (actual.contains(testdata)) {
					logger.error("Text found which was not supposed to be present!");
					logger.error("Text found was " + actual);
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertNotEquals(actual, testdata,
							"Text expected: '" + testdata + "', actual :'" + actual + "'");
				}
			} else {
				//TODO implement strArrayBuffer logging
				boolean isValidated = false;
				String expectedStr = null;
				for (String str : strArrayBuffer) {
					if (expectedStr == null)
						expectedStr = str;
					else
						expectedStr = expectedStr + ", " + str;
					if (actual.equalsIgnoreCase(str)) {
						isValidated = true;
						// validatedStr = str;
						break;
					}
				}
				if (isValidated) {
					//status = "Fail";
					//comment = "Expected one of '" + expectedStr + "', actual '" + actual + "'";
					// LogStep();
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					Assert.assertNotEquals(actual, expectedStr,
							"Expected one of '" + expectedStr + "', actual '" + actual + "'");
				} else {
					//comment = "Element text found is '" + actual + "'";
				}
			}
		}
		if (!isFound) {
			logger.warn("Object " + objectid + " not found.  This may or may not be intended page behavior");
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	public static void verifyTextNotPresentFromAD(String identifier, String testName, String step, String idType, String textToVerify, WebDriver driver) {

		logger.info("Verifying " + identifier + " is not present on the page from TCO");

		WebElement object = locateElementSwitch(idType, identifier, driver);

		if (object != null) {
			String actual = object.getText();
			if (actual.contains(textToVerify)) {
				logger.error("Text found which was not supposed to be present!");
				logger.error("Text found was " + actual);
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				Assert.fail("Text was not expected: '" + textToVerify + "', actual: '" + actual + "'");
			}
		}
		else{
			logger.warn("Object " + identifier + " not found.  This may or may not be intended page behavior");
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	/**
	 * This method takes in a String representing a label on the Excel data
	 * sheet. This method verifies that the text found on the page is the color
	 * found in the Excel data sheet. The color format is 'rgba(0, 0, 0, 1)'
	 * 
	 * @param field_label
	 *            a String representing a label on the Excel data sheet
	 * @param object
	 */
	public static void verifyTextColor(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying the color of " + field_label + " on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			String actual = object.getCssValue("color");

			if (!actual.contains(testdata)) {
				logger.error("Text had an unexpected color!");
				logger.error("Expected text color: " + testdata);
				logger.error("  Actual text color: " + actual);
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				Assert.assertEquals(actual, testdata,
						"Text color expected: '" + testdata + "', actual :'" + actual + "'");
			}
			logger.info("Expected text color: " + testdata);
			logger.info("  Actual text color: " + actual);
		}
		if (!isFound) {
			logger.error("Object " + objectid + " not found and its color could not be verified");
			Assert.fail("Object " + objectid + " was not found and its color could not be verified");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	/**
	 * This method takes in a String representing a label on the Excel data
	 * sheet This method verifies that the object pulled from the data sheet
	 * does not exist on the page.
	 * 
	 * @param field_label
	 *            a String representing a label on the Excel data sheet
	 * @param object 
	 */
	public static void verifyNotExists(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying " + field_label + " not on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		String objectID = strArray[1];
		String identifier = strArray[2];

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}

		WebElement object = allCheck(objectID, identifier, driver);
		if(object != null){
			logger.error("Object " + objectID + " should not be present!");
		}
		Assert.assertFalse(object != null, "Element should not exist on this page.");
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method finds elements on screen by tags and verifies that they contain
	 * the text pulled from the Excel sheet.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 * @return Returns true if the elements are verified, and false if not.
	 */
	public static boolean verifyTags(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying that objects exist on the page based on tags");
		logger.info("Field label is " + field_label);

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			List<WebElement> rows = object.findElements(By.tagName("label"));

			boolean isFound = false;
			for(Iterator<WebElement> i = rows.iterator(); i.hasNext();){
				String actual = i.next().getText();
				logger.debug("actual = " + actual);
				if(actual.contains(testdata)){
					if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
						captureScreenShot(driver, testName, step);
					}
					return true;
				}
			}
			if (!isFound) {
				logger.error("Element text expected '" + testdata + "' is not found");
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				Assert.fail("Element " + objectid + " was not found or verified");
				return false;
			} else {
				logger.error("verifyTags broke!!!");
				return false;
			}
		}

		else {
			logger.error("object " + objectid + " not found");
			Assert.fail("Object " + objectid + " was not found on the page");
			return false;
		}

	}

	public static void verifyNotInDropdown(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying that elements are not in the dropdown " + field_label);

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		String objID = strArray[1];
		String identifier = strArray[2];

		ArrayList<String> dropdownList = strArrayBuffer;
		String notExpected = null;

		logger.trace("Looping through the list from Excel to build a string containing values not expected");
		for (String str : dropdownList) {
			if (notExpected == null) {
				notExpected = str;
			} else
				notExpected = notExpected + ", " + str;
		}

		WebElement object = allCheck(objID, identifier, driver);
		if (object != null) {
			logger.trace("Creating a Select element from the object found on the page and grabbing the options");
			Select dropDown = new Select(object);
			List<WebElement> Options = dropDown.getOptions();

			logger.trace("Looping through the list from Excel");

			List<String> foundOptions = new ArrayList<String>();

			for(Iterator<String> i = dropdownList.iterator(); i.hasNext();){
				String str = i.next();
				logger.trace("Looping through the dropdown list from the page");
				for(Iterator<WebElement> i2 = Options.iterator(); i2.hasNext();){
					String option = i2.next().getText();
					logger.trace("Comparing " + str + " to " + option);
					if(option.equals(str)){
						foundOptions.add(str);
					}
				}
			}

			//			for (String str : dropdownList) {
			//				boolean isThere = false;
			//				logger.trace("Looping through the dropdown list from the page");
			//				for (WebElement option : Options) {
			//					logger.trace("Comparing " + str + " to " + option.getText());
			//					if (option.getText().equals(str)) {
			//						if (optionsString == null) {
			//							optionsString = str;
			//						} else {
			//							optionsString = optionsString + ", " + str;
			//						}
			//						logger.warn("Option found but was not supposed to be present: " + option);
			//						isThere = true;
			//						break;
			//					}
			//				}
			//				if (isThere) {
			//					if (found == null) {
			//						found = str;
			//					} else
			//						found = found + ", " + str;
			//				}
			//			}
			if (foundOptions.size() > 0) {
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				logger.error("Objects found in dropdown that were not supposed to be present:");
				for(String option : foundOptions){
					logger.error(option);
				}
				Assert.fail("Elements were found in dropdown that were not supposed to be present");
			} else {
				logger.debug("All elements confirmed to not be present");
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
			}
		}
	}

	public static void verifyTagsNotPresent(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying that tags are not present on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			List<WebElement> rows = object.findElements(By.tagName("label"));
			List<String> foundTags = new ArrayList<String>();

			logger.trace("Looping through the tags found on the page");
			for (Iterator<WebElement> i = rows.iterator(); i.hasNext();) {
				String actual = i.next().getText();
				//System.out.println("actual = " + actual);
				if (actual.contains(testdata)) {
					foundTags.add(actual);
				}
			}
			if (foundTags.size() > 0) {
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				logger.error("Unexpected tags found");
				for(String tag : foundTags){
					logger.error(tag);
				}
				Assert.fail("Tags found that were not supposed to be present on the page");
			} else {
				logger.error("verifyTags broke!!!");
			}
		}

		else {
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			logger.error("object " + objectid + " not found");
			Assert.fail("Object not found on the page, could not confirm if tags were present or not");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method verifies that the dropdown list contains the selections pulled
	 * from the Excel sheet.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void verifyDropdownList(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying elements in a dropdown or list");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String objectid = str_array[1];
		String identifier = str_array[2];

		List<String> dropdownList = strArrayBuffer;

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			Select dropDown = new Select(object);
			List<WebElement> Options = dropDown.getOptions();

			logger.trace("Looping through expected options");
			for (Iterator<String> i = dropdownList.iterator(); i.hasNext();) {
				String iText = i.next();
				boolean foundEle = false;
				logger.trace("Looping through options in the dropdown or list");
				for (Iterator<WebElement> i2 = Options.iterator(); i2.hasNext();) {
					String i2Text = i2.next().getText();
					logger.trace("Comparing " + iText + " to " + i2Text);
					if (i2Text.equals(iText)) {
						i2.remove();
						foundEle = true;
						break;
					}
				}
				if(foundEle){
					i.remove();
				}
			}
			if (!dropdownList.isEmpty()) {
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				logger.error("Expected options not found: ");
				for(String option : dropdownList){
					logger.error(option);
				}
				Assert.fail("Elements in dropdown not found.");
			} else {
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				logger.debug("All expected options found");
			}
		}
		if (!isFound) {
			logger.error("object " + objectid + " not found");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.fail("Element not found, could not be verified.");
		}
	}

	public static void verifyPageLabels(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath){

		logger.info("Verifying the labels specified in the Excel sheet");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String objectid = str_array[1];
		String identifier = str_array[2];

		logger.trace("Setting the labels to check to a list");
		List<String> labelsToCheck = new ArrayList<String>();
		labelsToCheck = strArrayBuffer;
		strArrayBuffer = null;

		WebElement object = allCheck(objectid, identifier, driver);
		List<String> labelArray = new ArrayList<String>();

		logger.trace("Building an arrayList of the labels inside the object specified");
		//build an array of the labels inside of the object specified in the Excel sheet
		for(WebElement ele : object.findElements(By.xpath("//td[contains(@class, 'labelCol')]"))){
			if(!ele.getText().isEmpty() && !ele.getText().equals("") && !ele.getText().equals(" ")){
				labelArray.add(ele.getText());
			}
		}
		for(WebElement ele : object.findElements(By.xpath("//label"))){
			if(!ele.getText().isEmpty() && !ele.getText().equals("") && !ele.getText().equals(" ") && !labelArray.contains(ele.getText())){
				labelArray.add(ele.getText());
			}
		}

		for(Iterator<String> iterator = labelsToCheck.iterator(); iterator.hasNext();){
			String labelToCheck = iterator.next();
			logger.trace("Checking " + labelToCheck);
			boolean found = false;
			logger.trace("Starting to loop through the labels from the page");

			for(Iterator<String> iterator2 = labelArray.iterator(); iterator2.hasNext();){
				String label = iterator2.next();
				logger.trace("Comparing " + labelToCheck + " to " + label);
				if(label.contains(labelToCheck)){
					logger.trace("Found " + labelToCheck + " in the list of labels on the page");
					found = true;
					logger.trace("Removing " + label + " from the list of labels on the page");
					iterator2.remove();
					break;
				}
			}
			if(found == true){
				logger.trace("Removing " + labelToCheck + " from the list of labels to check");
				iterator.remove();
			}
		}
		if(!labelsToCheck.isEmpty()){
			logger.error("Not all labels found and verified, there are still labels not found");
			for(String label : labelsToCheck){
				logger.warn("Label not found: " + label);
			}
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
		}
		
		if(!labelArray.isEmpty()){
			logger.debug("Leftover labels on the page:");
			for(String label : labelArray){
				logger.debug(label);
			}
		}
		Assert.assertTrue(labelsToCheck.isEmpty(), "Not all labels were found on the page!");
		logger.debug("All labels found and verified");
	}

	public static void verifyTableValues(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath){
		logger.info("Verifying the table values specified in the Excel sheet");
		String sArg = null;

		logger.trace("reading the Excel sheet");
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String objectid = str_array[1];
		String identifier = str_array[2];

		logger.trace("Setting the values to check to a list");
		List<String> dataToCheck = new ArrayList<String>();
		dataToCheck = strArrayBuffer;
		strArrayBuffer = null;

		WebElement object = allCheck(objectid, identifier, driver);
		List<String> tableArray = new ArrayList<String>();

		logger.trace("Building an arrayList of the labels inside the object specified");
		//build an array of the labels inside of the object specified in the Excel sheet
		for(WebElement ele : object.findElements(By.tagName("td"))){
			tableArray.add(ele.getText());
		}
		for(WebElement ele : object.findElements(By.tagName("th"))){
			tableArray.add(ele.getText());
		}

		logger.trace("dataToCheck's size is " + dataToCheck.size());
		logger.trace("tableArray's size is " + tableArray.size());

		for(Iterator<String> iterator = dataToCheck.iterator(); iterator.hasNext();){
			String datumToCheck = iterator.next();
			logger.trace("Checking " + datumToCheck);
			boolean found = false;
			logger.trace("Starting to loop through the labels from the page");

			for(Iterator<String> iterator2 = tableArray.iterator(); iterator2.hasNext();){
				String cell = iterator2.next();
				logger.trace("Comparing " + datumToCheck + " to " + cell);
				if(datumToCheck.equals(cell)){
					logger.debug("Found " + datumToCheck + " in the list of cells on the page");
					found = true;
					logger.trace("Removing " + cell + " from the list of cells on the page");
					iterator2.remove();
					break;
				}
			}
			if(found == true){
				logger.trace("Removing " + datumToCheck + " from the list of data to check");
				iterator.remove();
			}
		}
		if(!dataToCheck.isEmpty()){
			logger.error("Not all labels found and verified, there are still data not found");
			for(String label : dataToCheck){
				logger.error("datum not found: " + label);
			}
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
		}
		Assert.assertTrue(dataToCheck.isEmpty(), "Not all data were found on the page!");
		logger.info("All data found and verified");
	}

	public static void verifyNotClickable(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Verifying that a button is not clickable");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			if (!object.getAttribute("class").equals("btnDisabled")) {
				logger.error("Button is not disabled");
				if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
					captureScreenShot(driver, testName, step);
				}
				Assert.assertEquals(object.getAttribute("class"), "btnDisabled",
						"The button's class should be 'btnDisabled'");
			}
			logger.debug("Button found and is disabled.");
		}
		if (!isFound) {
			logger.error("Button not found");
			if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
				captureScreenShot(driver, testName, step);
			}
			Assert.fail("Button not found, could not be verified to be disabled.");
		}

		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
	}

	public static void verifyMFTooltip(String field_label, WebDriver driver, String tabLabel, String excelPath) // TODO fully implement 
	{

		logger.info("Verifying that the MerchantForce tooltip has the correct text");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] strArray = sArg.split(";");
		String objectid = strArray[1];
		String testData = strArray[0];
		String identifier = strArray[2];

		Actions action = new Actions(driver);

		WebElement object = allCheck(objectid, identifier, driver);

		action.moveToElement(object);
		action.perform();

		WebElement helpText = driver.findElement(By.className("helpText"));
		Assert.assertEquals(helpText.getText(), testData, "Help text was incorrect or not found.");
	}

	public static void selectDropdownHardcoded(String id, String idType, String value, WebDriver driver){
		logger.info("Selecting " + value + " from " + id + " hardcoded");
		WebElement dropdownObj = allCheck(id, idType, driver);
		if(dropdownObj != null){
			logger.debug("Found " + id + ", now selecting " + value);
			Select dropdownSel = new Select(dropdownObj);
			dropdownSel.selectByVisibleText(value);
		}
		else{
			logger.warn("Object " + id + " not found, could not select " + value);
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method selects one or more elements from a dropdown list, multiselect
	 * list, or list.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void selectFromList(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Selecting one or more options in a dropdown or list");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] strArray = sArg.split(";");
		String objectid = strArray[1];
		String testData = strArray[0];
		String identifier = strArray[2];

		boolean isFound = false;

		//TODO write logging for strArrayBuffer
		if (testData.contains("strArrayBuffer")) {
			ArrayList<String> selectList = strArrayBuffer;

			WebElement object = allCheck(objectid, identifier, driver);
			if (object != null) {
				isFound = true;
				Select selectListObj = new Select(object);
				List<WebElement> options = selectListObj.getOptions();

				for (String str : selectList) {
					for (WebElement ele : options) {
						if (ele.getText().equals(str)) {
							ele.click();
							break;
						}
					}
				}
			}
			if (!isFound) {
				logger.warn("object " + objectid + " not found");
			}
		} else {
			WebElement object = allCheck(objectid, identifier, driver);
			if (object != null) {
				isFound = true;
				logger.trace("Creating a Select from the found object");
				Select dropDown = new Select(object);
				try {
					dropDown.deselectAll();
					logger.trace("Deselected all options");
				} catch (UnsupportedOperationException e) {
					logger.debug("Object was not a multiselect, could not deselect options");
				}

				dropDown.selectByVisibleText(testData);
			}
			if (!isFound) {
				logger.warn("object " + objectid + " not found");
			}
		}
	}

	public static void selectFromList_WithLabel(String field_label, String testdata, WebDriver driver) {

		logger.info("Selecting from a list or dropdown using the object's label");

		boolean isFound = false;
		WebElement object = null;
		object = elementExists(By.xpath("id(//label[text() = '" + field_label + "']/@for)"), driver);
		if (object != null) {
			isFound = true;
			Select dropDown = new Select(object);
			dropDown.selectByVisibleText(testdata);
		}
		if (!isFound) {
			logger.warn("object " + field_label + " not found");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method selects an element from a dropdown list.
	 * 
	 * @deprecated use selectFromList if possible. This function still works,
	 *             but is less flexible.
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object
	 */
	public static void selectFromDropdown(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.warn("Using deprecated function 'selectFromDropdown'");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];

		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			Select dropDown = new Select(object);
			dropDown.selectByVisibleText(testdata);
		} else {
			System.out.println("object " + objectid + " not found");
		}
	}

	public static void toggleCheckbox(String field_label, boolean onOrOff, WebDriver driver, String tabLabel, String excelPath) {

		String toggleTo;
		if(onOrOff == true){
			toggleTo = "On";
		} else {
			toggleTo = "Off";
		}
		logger.info("Toggling checkbox to " + toggleTo);

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] strArray = sArg.split(";");
		String str = strArray[1];
		String identifier = strArray[2];

		boolean isFound = false;
		WebElement object = allCheck(str, identifier, driver);
		if (object != null) {
			isFound = true;
			if (object.isDisplayed()) {
				if (object.isSelected() != onOrOff) {
					object.click();
				}
			} else {
				logger.warn("Element was not visible");
			}
		}
		if (!isFound) {
			logger.warn("object " + str + " not found");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method clicks on an element.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object
	 */
	public static void click_Button_Radio_Checkbox(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Clicking an object on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		String str = strArray[1];
		String identifier = strArray[2];
		WebElement object = allCheck(str, identifier, driver);
		if(object != null){
			object.click();
		}
		else{
			logger.warn("Object " + str + " not found");
		}
	}

	public static void clickFromAD(String identifier, String idType, WebDriver driver) {

		logger.info("Clicking an object on the page with info from TCO");

		WebElement object = locateElementSwitch(idType, identifier, driver);
		if(object != null){
			object.click();
		}
		else{
			logger.warn("Object " + identifier + " not found");
		}
	}

	public static void click_Button_Radio_Checkbox_WithValue(String field_value, WebDriver driver) {

		logger.info("Clicking an object on the page using a value");

		boolean isFound = false;
		WebElement object = elementExists(By.xpath("//*[@value='" + field_value + "')]"), driver);
		if (object != null) {
			isFound = true;
			if (object.isDisplayed()) {
				object.click();
			} else {
				logger.warn("Element was not visible");
			}
		}
		if (!isFound) {
			logger.warn("object " + field_value + " not found");
		}
	}

	public static void click_Button_Radio_Checkbox_WithLabel(String field_label, WebDriver driver) {

		logger.info("Clicking an object using its label");

		boolean isFound = false;
		WebElement object = elementExists(By.xpath("id(//label[text() = '" + field_label + "']/@for)"), driver);
		if (object != null) {
			isFound = true;
			if (object.isDisplayed()) {
				object.click();
			} else {
				logger.warn("Element was not visible");
			}
		}
		if (!isFound) {
			logger.warn("object " + field_label + " not found");
		}
	}

	/**
	 * This method takes in a String representing a label on the Excel data
	 * sheet. This method double clicks on an element on the page.
	 * 
	 * @param field_label
	 *            a String representing a label on the Excel data sheet
	 * @param object 
	 */
	public static void doubleClick_object(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Double-clicking an object on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] strArray = sArg.split(";");
		String str = strArray[1];
		String identifier = strArray[2];

		boolean isFound = false;
		WebElement object = allCheck(str, identifier, driver);
		if (object != null) {
			isFound = true;
			if (object.isDisplayed()) {
				Actions act = new Actions(driver);
				logger.trace("Moving to and double-clicking " + str);
				act.moveToElement(object).doubleClick().build().perform();
			} else {
				logger.warn("Element was not visible");
			}
		}
		if (!isFound) {
			logger.warn("object " + str + " not found");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet, and an
	 * int representing how long the random suffix should be. This method enters
	 * a base String pulled from the Excel data sheet followed by a random
	 * string into an editbox. The random String's length is based on the int
	 * passed into the method.
	 * 
	 * @param base
	 *            A String representing a label on the Excel data sheet.
	 * @param suffixLength
	 *            The length of the random string to be appended to the base
	 *            String.
	 * @param object
	 */
	public static void suffixBase(String base, int suffixLength, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Creating a random suffix for a data value");

		String sArg = null;
		sArg = readExcel(base, tabLabel, excelPath);
		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		String objectid = str_array[1];

		double random = Math.random();
		for (int i = 1; i <= suffixLength; i++) {
			random = random * 10;
		}
		int random1 = (int) random;
		String Return_Data = testdata + random1 + ";" + objectid;
		fillEditbox(Return_Data, driver, tabLabel, excelPath);
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method enters a String pulled from the Excel data sheet into an editbox.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void fillEditbox(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Using " + field_label +" to fill in a text field");

		String sArg = field_label;
		if (!field_label.contains(";")) {
			sArg = readExcel(field_label, tabLabel, excelPath);
		}

		String[] str_array = sArg.split(";");
		String testdata = str_array[0];
		logger.trace("testdata contains " + testdata);
		String objectid = str_array[1];
		logger.trace("objectid contains " + objectid);
		String identifier = str_array[2];
		logger.trace("identifier contains " + identifier);

		WebElement object = allCheck(objectid, identifier, driver);
		if(object != null)
		{
			object.clear();
			object.sendKeys(testdata);
		}
		else
		{
			logger.warn("Text box not found, cannot enter text.");
		}
	}

	public static void fillEditBoxFromAD(String identifier, String idType, String textToEnter, WebDriver driver) {

		logger.info("Filling a text field with data from TCO");

		WebElement object = locateElementSwitch(idType, identifier, driver);
		object.clear();
		object.sendKeys(textToEnter);
	}

	public static void fillEditbox_WithLabel(String field_label, String testdata, WebDriver driver) {

		logger.info("Filling a text field using its label");

		boolean isFound = false;
		int count = 0;
		do {
			WebElement object = elementExists(By.xpath("id(//label[text() = '" + field_label + "']/@for)"), driver);
			if (object != null) {
				isFound = true;
				object.clear();
				object.sendKeys(testdata);
			}
			count++;
		} while (!isFound && count < 3);
		if (!isFound) {
			logger.warn("object " + field_label + " not found");
		}
	}

	/**
	 * This function takes a String that represents a column on the Excel data
	 * sheet. This function sets the SimpleDateFormat based off of the data in
	 * the column if found.
	 * 
	 * @param field_label
	 */
	public static void setSDF(String field_label, String tabLabel, String excelPath) {

		logger.info("Setting the SimpleDateFormat");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		dateFormat = strArray[0];
		sdf.applyPattern(dateFormat);
	}

	/**
	 * This function takes a String that represents a column on the Excel data
	 * sheet. This function sets the SimpleDateFormat based off of the data in
	 * the column if found.
	 * 
	 * @param newDateFormat
	 */
	public static void setDateTimeFormat(String newDateFormat, String tabLabel, String excelPath) {

		logger.info("Setting the date/time format");

		String sArg = null;
		sArg = readExcel(newDateFormat, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		newDateFormat = strArray[0];
		sdf.applyPattern(newDateFormat);
	}

	public static void captureDateTime(String captureField, String tabLabel, String excelPath) {

		logger.info("Capturing the date/time to the datasheet");

		String sArg = null;
		sArg = readExcel(captureField, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		int colLoc = Integer.parseInt(str_array[0]);

		try {
			File file = new File(excelPath);
			FileInputStream inputStream = new FileInputStream(file);
			XSSFWorkbook databook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = databook.getSheet(tabLabel);
			XSSFRow row = sheet.getRow(0);
			XSSFCell cell = row.getCell(colLoc);
			XSSFRow datarow = sheet.getRow(2);
			XSSFCell datacell = datarow.getCell(colLoc - 1);

			if (datacell == null) {
				datacell = datarow.createCell(colLoc - 1);
				datacell.setCellType(XSSFCell.CELL_TYPE_STRING);
			}

			GetTimeStampValue();
			String dateTimeToStore = sdf.format(time);

			datacell.setCellValue(dateTimeToStore);

			cell.setCellValue("all");
			FileOutputStream outputStream = new FileOutputStream(file);
			databook.write(outputStream);
			outputStream.close();
			databook.close();

		} catch (IOException e) {
			e.printStackTrace();
			logger.error("The Excel file was not able to be written to");
			Assert.fail("Excel file was not able to be written to.", e);
		}
	}

	public static void setColType(String field_label, String newType,String excelPath, String tabLabel) {

		logger.info("Setting the Excel datasheet column type");

		try {
			File file = new File(excelPath);
			FileInputStream inputStream = new FileInputStream(file);
			XSSFWorkbook databook = new XSSFWorkbook(inputStream);
			XSSFSheet readsheet = databook.getSheet(tabLabel);
			XSSFRow labelrow = readsheet.getRow(0);
			int noOfColumns = readsheet.getRow(0).getPhysicalNumberOfCells();
			databook.getCreationHelper().createFormulaEvaluator().evaluateAll();

			boolean isFound = false;

			for (int i = 1; i < noOfColumns; i++) {
				String labelHeader = labelrow.getCell(i).getStringCellValue();
				if (field_label.equalsIgnoreCase(labelHeader)) {
					isFound = true;
					labelrow.getCell(i + 1).setCellValue(newType);
					FileOutputStream outputStream = new FileOutputStream(file);
					databook.write(outputStream);
					outputStream.close();
					break;
				}
			}
			if (!isFound) {
				logger.error("The label header could not be found!");
				Assert.fail("Label header '" + field_label + "' was not found in data sheet '" + tabLabel + "'");
			}
			databook.close();
		} catch (IOException e) {
			Assert.fail("Excel File was unable to be interacted with");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method captures a field on the page and stores it in the Excel data sheet
	 * for later use. It changes the label's type from 'capture' to 'single'
	 * 
	 * @param captureField
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void captureData(String captureField, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Capturing data to the Excel datasheet");

		String sArg = null;
		sArg = readExcel(captureField, tabLabel, excelPath);

		String[] str_array = sArg.split(";");
		int colLoc = Integer.parseInt(str_array[0]);
		String objectid = str_array[1];
		String identifier = str_array[2];

		try {
			File file = new File(excelPath);
			FileInputStream inputStream = new FileInputStream(file);
			XSSFWorkbook databook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = databook.getSheet(tabLabel);
			XSSFRow row = sheet.getRow(0);
			XSSFCell cell = row.getCell(colLoc);
			XSSFRow datarow = sheet.getRow(2);
			XSSFCell datacell = datarow.getCell(colLoc - 1);

			if (datacell == null) {
				datacell = datarow.createCell(colLoc - 1);
				datacell.setCellType(XSSFCell.CELL_TYPE_STRING);
			}

			boolean isFound = false;
			WebElement object = allCheck(objectid, identifier, driver);
			if (object != null) {
				isFound = true;
				String dataToStore = object.getText();
				datacell.setCellValue(dataToStore);
			}
			if (!isFound) {
				logger.error("object " + objectid + " not found");
				Assert.fail("Object " + objectid + " not found");
			}
			cell.setCellValue("single");
			FileOutputStream outputStream = new FileOutputStream(file);
			databook.write(outputStream);
			outputStream.close();
			databook.close();
		} catch (IOException e) {
			logger.error("The Excel file was unable to be written to");
			Assert.fail("Excel file was not able to be written to.", e);
		}
	}

	/**
	 * Takes in a String of the name of who should be logged in. If logged in to
	 * MerchantForce, this method logs in as another user.
	 * 
	 * @param field_label
	 *            A String name.
	 * @param object
	 */
	public static String loginUser(String field_label, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Logging in as another User in SalesForce");

		waitForLoad(driver);

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);

		String[] strArray = sArg.split(";");
		String testdata = strArray[0];
		String currentUser = testdata;
		// String objectid = strArray[1];

		// Click on setup link

		// boolean setupFound = false;
		// int count = 0;
		// do
		// {
		// if (elementExists(By.id("setupLink")))
		// {
		// object.click();
		// setupFound = true;
		// }
		// try
		// {
		// //Thread.sleep(500);
		// }
		// catch (InterruptedException e)
		// {
		// e.printStackTrace();
		// }
		// count++;
		// }
		// while(!setupFound && count < 3);
		// if(!setupFound)
		// {
		// status = "Fail";
		// comment = "Object 'Setup Link' not found, tried " + count + "
		// times.";
		// System.out.println("Navigation to Setup failed");
		// }

		// Enter Switch To User's name in Editbox

		boolean userNavLabelFound = false;

		// opening the navigation dropdown
		WebElement object = elementExists(By.id("userNavLabel"), driver);
		if (object != null) {
			userNavLabelFound = true;
			object.click();
		}
		if (!userNavLabelFound) {
			logger.error("User navigation label not found");
			Assert.fail("User navigation label was not found, failing test");
		}

		boolean setupFound = false;

		// clicking Setup
		object = elementExists(By.linkText("Setup"), driver);
		if (object != null) {
			setupFound = true;
			object.click();
		}
		if (!setupFound) {
			logger.error("Setup button not found");
			Assert.fail("Setup button not found, failing test");
		}

		waitForLoad(driver);

		boolean searchFound = false;

		// using the search in the Setup page
		object = elementExists(By.id("setupSearch"), driver);
		if (object != null) {
			object.clear();
			object.sendKeys(testdata);
			object.sendKeys(Keys.ENTER);
			searchFound = true;
		}
		if (!searchFound) {
			logger.error("Search box not found.");
			Assert.fail("Setup search box not found, failing test");
		}

		// click on User Name

		waitForLoad(driver);

		boolean userFound = false;

		object = elementExists(By.linkText(testdata), driver);
		if (object != null) {
			object.click();
			userFound = true;
		}
		if (!userFound) {
			logger.error("User not found or clicked on.");
			Assert.fail("User was not found or clicked on, failing test");
		}

		waitForLoad(driver);

		// Click on Login

		boolean loginButtonFound = false;

		object = elementExists(By.name("login"), driver);
		if (object != null) {
			object.click();
			loginButtonFound = true;
		} else {
		}

		waitForLoad(driver);

		WebDriverWait wait = new WebDriverWait(driver, 45);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@title='Alert Tab - Selected']")));

		waitForLoad(driver);

		if (!loginButtonFound) {
			logger.error("User not successfully logged in- login button not found.");
			Assert.fail("User not sucessfully logged in, failing test");
		}
		return currentUser;
	}

	/**
	 * This method logs the current user out. If logged in as someone else, logs
	 * that person out. If not logged in as someone else, just logs out.
	 * @param object 
	 */
	public static void logoutUser(WebDriver driver) {
		// Click on User Navigation button and Logout

		logger.info("Logging out of current User in SalesForce");

		boolean navLabelFound = false;

		// Open the navigation dropdown
		WebElement object = elementExists(By.id("userNavLabel"), driver);
		if (object != null) {
			navLabelFound = true;
			object.click();
		}

		if (!navLabelFound) {
			logger.error("User navigation was not opened.");
		}

		boolean logoutFound = false;

		// click Logout
		object = elementExists(By.linkText("Logout"), driver);
		if (object != null) {
			logoutFound = true;
			object.click();
		}

		if (!logoutFound) {
			logger.error("User was not successfully logged out.");
		}
	}

	/**
	 * Logs out of the current user and into a different user. Only works if
	 * already logged in as someone else.
	 * 
	 * @param switchToUser
	 *            User to switch to.
	 * @param object
	 */
	public static String switchUser(String switchToUser, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Switching from one User to another User in SalesForce");

		logoutUser(driver);
		waitForLoad(driver);
		return loginUser(switchToUser, driver, tabLabel, excelPath);
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. Attaches
	 * a file to the file browser.
	 * 
	 * @param field_label
	 *            A string representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void attachFile(String field_label, WebDriver driver, String tabLabel, String excelPath, String testName) {

		logger.info("Attaching a file");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] str_array = sArg.split(";");
		String file = str_array[0];
		String objectid = str_array[1];
		String identifier = str_array[2];
		String fileOutLoc;

		if (!SystemUtils.IS_OS_MAC) {
			ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
			InputStream is = classLoader.getResourceAsStream(file);

			fileOutLoc = System.getProperty("user.dir") + File.separator + file;
			File fileOut = new File(fileOutLoc);
			fileOut.deleteOnExit();

			try{
				FileOutputStream outputStream = new FileOutputStream(fileOut);
				IOUtils.copy(is, outputStream);
				outputStream.close();
				is.close();
			}catch(IOException e){
				logger.error("The file that failed to be created was supposed to be " + fileOutLoc);
			}
		}
		// if running on Mac
		else {
			ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
			InputStream is = classLoader.getResourceAsStream(file);

			fileOutLoc = "/Users/" + System.getProperty("user.name") + "/Documents/" + file;
			File fileOut = new File(fileOutLoc);
			fileOut.deleteOnExit();

			try{
				FileOutputStream outputStream = new FileOutputStream(fileOut);
				IOUtils.copy(is, outputStream);
				outputStream.close();
				is.close();
			}catch(IOException e){
				logger.error("IO Exception occurred");
			}
		}
		LocalFileDetector detector = null;
		File f = null;
		if(System.getProperty("remoteRun").equalsIgnoreCase("true")){
			detector = new LocalFileDetector();
			f = detector.getLocalFile(fileOutLoc);
		}

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
			if(System.getProperty("remoteRun").equalsIgnoreCase("true")){
				((RemoteWebElement)object).setFileDetector(detector);
				String pathToSend = f.getAbsolutePath();
				object.sendKeys(pathToSend);
			}
			else{
				object.sendKeys(fileOutLoc);
			}
		}
		if (!isFound) {
			logger.warn("object " + objectid + " not found");
		}
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. Asserts
	 * that an element on the page exists. If it does not exist, the test run is
	 * ended and the test is failed.
	 * 
	 * @param field_label
	 *            A String representing a label on the Excel data sheet.
	 * @param object 
	 */
	public static void assertElementExists(String field_label, String testName, String step, WebDriver driver, String tabLabel, String excelPath) {

		logger.info("Asserting that an element exists on the page");

		String sArg = null;
		sArg = readExcel(field_label, tabLabel, excelPath);
		String[] str_array = sArg.split(";");
		String objectid = str_array[1];
		String identifier = str_array[2];

		boolean isFound = false;
		WebElement object = allCheck(objectid, identifier, driver);
		if (object != null) {
			isFound = true;
		}
		if(System.getProperty("remoteRun").equalsIgnoreCase("false")){
			captureScreenShot(driver, testName, step);
		}
		Assert.assertTrue(isFound, "Object " + objectid + "was not found.");
	}

	/**
	 * Takes in a String representing a label on the Excel data sheet. This
	 * method reads the selected label in the Excel data sheet. Depending on
	 * what the column is set to, this method will return the id of the element
	 * to interact with, as well as any data that needs to be used by the method
	 * calling this one.
	 * 
	 * @param passed_label
	 *            A String representing a label on the Excel data sheet.
	 * @return A String containing the id of the object to interact with, as
	 *         well as any data the calling function might need, separated by
	 *         ';'
	 */
	public static String readExcel(String passed_label, String tabLabel, String xlpath) {

		logger.info("Pulling data from the Excel sheet");

		logger.trace("Passed label is " + passed_label);
		logger.trace("Tab label is " + tabLabel);
		logger.trace("Excel path is " + xlpath);

		String testData;
		String usedStatus;
		String objID = null;
		String returnData = null;
		String identifier = "";

		try{
			File file = new File(xlpath);
			FileInputStream inputStream = new FileInputStream(file);
			XSSFWorkbook databook = new XSSFWorkbook(inputStream);
			XSSFSheet readsheet = databook.getSheet(tabLabel);
			XSSFRow labelrow = readsheet.getRow(0);
			XSSFRow objidrow = readsheet.getRow(1);
			int noOfColumns = readsheet.getRow(0).getPhysicalNumberOfCells();
			int noOfRows = readsheet.getLastRowNum();
			databook.getCreationHelper().createFormulaEvaluator().evaluateAll();

			for (int i = 1; i < noOfColumns; i = i + 2) {
				String labelHeader = labelrow.getCell(i).getStringCellValue();
				logger.trace("Label header being checked is " + labelHeader);
				// finding the label header
				if(labelHeader.equalsIgnoreCase("<label>")){
					returnData = "Label header not found";
					break;
				}
				if (passed_label.equalsIgnoreCase(labelHeader)) {
					logger.trace("Label found was " + labelrow.getCell(i + 1).getStringCellValue());
					try {
						if (objidrow.getCell(i + 1).getStringCellValue().equals("")) {
							identifier = "none";
						} else {
							identifier = objidrow.getCell(i + 1).getStringCellValue();
						}
					} catch (NullPointerException e) {
						identifier = "none";
					}
					// if the label type is 'any'
					String dataType = labelrow.getCell(i + 1).getStringCellValue();
					if (dataType.equalsIgnoreCase("any")) {
						logger.trace("Label type found was any");
						// looping through the rows of the sheet
						for (int j = 2; j < noOfRows; j++) {
							XSSFRow datarow = readsheet.getRow(j);
							usedStatus = datarow.getCell(i + 1).getStringCellValue();

							// finding the first 'not used' data
							if (usedStatus.equalsIgnoreCase("not used") || usedStatus.isEmpty() || usedStatus == null) {
								testData = datarow.getCell(i).getStringCellValue();
								objID = objidrow.getCell(i).getStringCellValue();

								Cell updateCell = readsheet.getRow(j).getCell(i + 1);
								updateCell.setCellValue("used");
								FileOutputStream outputStream = new FileOutputStream(file);
								databook.write(outputStream);
								outputStream.close();
								returnData = testData + ";" + objID;
								break;
							}
						}
						break;
					} else if (dataType.equalsIgnoreCase("all")) {
						logger.trace("Label type found was all");
						testData = "test data in strArrayBuffer";
						objID = objidrow.getCell(i).getStringCellValue();
						returnData = testData + ";" + objID;
						strArrayBuffer = new ArrayList<String>();

						int counter = 0;
						for (int j = 2; j <= noOfRows; j++) {
							XSSFRow curRow = readsheet.getRow(j);
							XSSFCell curCell = curRow.getCell(i);
							if (curCell != null && !curCell.getStringCellValue().equals("")) {
								counter++;
							} else
								break;
						}
						for (int j = 2; j <= counter + 1; j++) {
							String eleToAdd = readsheet.getRow(j).getCell(i).getStringCellValue();
							strArrayBuffer.add(j - 2, eleToAdd);
						}
						break;
					} else if (dataType.equalsIgnoreCase("single")) {
						logger.trace("Label type found was single");
						XSSFRow datarow = readsheet.getRow(2);
						testData = datarow.getCell(i).getStringCellValue();
						logger.trace("Test data is " + testData);
						if (objidrow.getCell(i) != null) {
							objID = objidrow.getCell(i).getStringCellValue();
						}
						returnData = testData + ";" + objID;
						logger.trace("Return data is " + returnData);

						break;
					} else if (dataType.equalsIgnoreCase("click")) {
						logger.trace("Label type found was click");
						objID = objidrow.getCell(i).getStringCellValue();
						returnData = "click;" + objID;

						break;
					} else if (dataType.equalsIgnoreCase("capture")) {
						logger.trace("Label type found was capture");
						String colLoc = (i + 1) + "";
						objID = objidrow.getCell(i).getStringCellValue();
						returnData = colLoc + ";" + objID;

						break;
					} else {
						logger.error("No proper label was found");
						logger.error("Label found was " + labelrow.getCell(i + 1).getStringCellValue());
					}
				} else {
					returnData = "Label header not found";
				}
			}
			databook.close();
			inputStream.close();
		} catch(IOException e){
			logger.warn("An IO exception was caught while trying to read the Excel file");
			logger.warn(e);
		}
		if (returnData != null) {
			if (returnData.equals("Label header not found")) {
				logger.error("The label header could not be found!");
				Assert.fail("Label header '" + passed_label + "' was not found in data sheet '" + tabLabel + "'");
			}
		} else
			logger.warn("No return data was set");
		if (identifier != null) {
			returnData = returnData + ";" + identifier;
		}

		return returnData;
	}

	// public static void enterCredentials(String username, String password)
	// {
	// WebDriverWait wait = new WebDriverWait(TMX_lib.driver, 10);
	// Alert alert = wait.until(ExpectedConditions.alertIsPresent());
	// alert.setCredentials(new UserAndPassword(username, password));
	// alert.accept();
	// driver.switchTo().defaultContent();
	// }

	/**
	 * Takes in a String message, an int for the top position of the message
	 * box, and an int for the left position of the message box.
	 * 
	 * @param msg
	 *            A String message to be displayed.
	 * @param top
	 *            The location that the top of the message box will be located
	 *            at.
	 * @param left
	 *            The location that the left of the message box will be located
	 *            at.
	 */
	public static void MsgBox(String msg, int top, int left) {
		final JOptionPane pane = new JOptionPane(msg);
		pane.setBackground(Color.red);
		final JDialog d = pane.createDialog((JFrame) null, "Alert");
		d.setLocation(top, left);
		d.setVisible(true);
		d.setAlwaysOnTop(true);
	}

	/**
	 * This method captures and saves a screenshot.
	 */
	public static void captureScreenShot(WebDriver driver, String testName, String step) {

		logger.debug("Capturing a screenshot");

		File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try{
			if (!SystemUtils.IS_OS_MAC) {
				new File("C:\\Users\\" +System.getProperty("user.name")
				+ "\\Documents\\Screenshots\\" + testName + "\\").mkdirs();
				FileUtils.moveFile(screenshotFile,
						new File("c:\\Users\\" + System.getProperty("user.name")
						+ "\\Documents\\Screenshots\\" + testName +"\\" + testName + "_Step" + step + "_"
						+ GetTimeStampValue() + ".png"));
				FileUtils.deleteQuietly(screenshotFile);
			} else {
				new File("/Users/" +System.getProperty("user.name")
				+ "/Documents/Screenshots/" + testName + "/").mkdirs();
				FileUtils.moveFile(screenshotFile,
						new File("/Users/" + System.getProperty("user.name")
						+ "/Documents/Screenshots/" + testName + "/" + testName + "_Step" + step + "_"
						+ GetTimeStampValue() + ".png"));
				FileUtils.deleteQuietly(screenshotFile);
			}
		}catch(IOException e){
			Assert.fail("Screenshot was unable to be captured/saved properly!");
		}
	}

	/**
	 * This method captures the current system time.
	 * 
	 * @return A String representing the current system time.
	 */
	public static String GetTimeStampValue() {

		logger.info("Getting the current date/time");

		cal = Calendar.getInstance();
		time = cal.getTime();
		String timestamp = time.toString();
		String systime = timestamp.replace(":", "-");
		return systime;
	}

	// /**
	// * This method creates a .csv format file to be used as a log for the
	// current test.
	// */
	// public static void LogCreate ()
	// {
	// // creates the run log at the beginning of the test
	//
	// try {
	// String sNow = new
	// SimpleDateFormat("MM-dd-yyyy_HH.mm").format(Calendar.getInstance().getTime());
	//
	// File file = new File (rlPath + testName + "_" + sNow + ".csv");
	// if (! file.exists()) file.createNewFile();
	//
	// rlOutFile = new FileWriter(file.getAbsoluteFile());
	// rlOutBW = new BufferedWriter(rlOutFile);
	//
	// //write first column header row
	// rlOutBW.write("Run Date,Test Name,Test
	// Step,Action,Narrative,Status,Comment");
	// rlOutBW.newLine();
	//
	// status = "Pass"; comment = ""; node = ""; description = ""; //reset
	// }
	// catch (FileNotFoundException e)
	// {
	// e.printStackTrace();
	// }
	// catch (IOException e)
	// {
	// e.printStackTrace();
	// }
	// }

	// /**
	// * This method records the information of the current test step in the log
	// created by LogCreate.
	// */
	// public static void LogStep ()
	// {
	// //Log Step code here
	// try
	// {
	// java.util.Date date= new java.util.Date();
	// String sNow = new SimpleDateFormat("MM-dd-yyyy HH:mm:ss").format(new
	// Timestamp(date.getTime()));
	//
	// //System.out.println(K + " " + step + " " + action + " " + narrative + "
	// " + status + " " + comment);
	//
	// if (status.equals("Fail")) failedTest = true; //flag any failed step as
	// failed test
	//
	// //write step level data row
	// rlOutBW.write(sNow+","+ testName +","+ step +","+ action +",\""+
	// narrative +"\","+ status +",\""+ comment + "\"");
	// rlOutBW.newLine();
	//
	// }
	// catch (Exception e)
	// {
	// System.out.println("Runlogging Error - " + e.getMessage());
	// }
	// }

	// /**
	// * This method closes the test log created by LogCreate.
	// */
	// public static void LogClose ()
	// {
	// // closes out the run log at the end of a test
	//
	// try
	// {
	// rlOutBW.close();
	// if (rlOutFile != null) rlOutFile.close();
	//
	// //log overall test result to external datasheet
	// File ofile = new File (dsPath + resultsDS);
	// boolean bNew = false;
	//
	// if (! ofile.exists())
	// {
	// ofile.createNewFile();
	// bNew = true;
	// }
	//
	// FileWriter resFile = new FileWriter(ofile.getAbsoluteFile(),true);
	// BufferedWriter resBW = new BufferedWriter(resFile);
	//
	// //write first column header row
	//
	// if (bNew)
	// {
	// resBW.write("Run Date,Test Name,Status");
	// resBW.newLine();
	// }
	//
	// java.util.Date date= new java.util.Date();
	// String stat = "";
	// if (failedTest) stat = "Failed";
	// else stat = "Passed";
	//
	// resBW.write(date +","+ testName + "," + stat);
	// resBW.newLine();
	// resBW.close();
	// resFile.close();
	// }
	// catch (FileNotFoundException e)
	// {
	// e.printStackTrace();
	// }
	// catch (IOException e)
	// {
	// e.printStackTrace();
	// }
	// }

	/**
	 * This method waits for the page to return a readystate of 'complete' then
	 * waits an additional second.
	 */
	public static void waitForLoad(WebDriver driver) {

		logger.trace("WaitForLoad triggered");

		WebDriverWait wait = new WebDriverWait(driver, 90);

		try {
			wait.withMessage("Page did not load")
			.until(ExpectedConditions.presenceOfElementLocated(By.xpath("html/body")));
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@src='/img/loading32.gif']")));
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@title='Alert Tab - Selected']")));
		} catch (UnhandledAlertException e) {
			logger.error("An unhandled alert appeared");
		}
		try {
			Thread.sleep(100);
		} catch (InterruptedException e) {
		}
	}

	/*public static void highlightElement(WebElement ele, WebDriver driver) {
		if (isFirefox || isChrome) {
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'border: solid 5px #BADA55')", ele);
		}
	}

	public static void highlightOff(WebElement ele, WebDriver driver) {
		if (isFirefox || isChrome) {
			JavascriptExecutor js = (JavascriptExecutor) driver;
			try {
				js.executeScript("arguments[0].setAttribute('style', 'border: initial')", ele);
			} catch (WebDriverException e) {
			}
		}
	}*/

	/**
	 * Takes in a By to try and find an element on the page. This method returns
	 * true if an element is found, false if not.
	 * 
	 * @param by
	 *            a locator to use to try and find an element on the page.
	 * @return True if element is found and displayed, false if not.
	 */
	public static WebElement elementExists(By by, WebDriver driver) {

		logger.trace("Searching for the element and setting it to the object to be manipulated");

		WebElement object;
		try{
			object = new WebDriverWait(driver, 10, 250).until(ExpectedConditions.presenceOfElementLocated(by));
			return object;
		}catch(TimeoutException t){
			object = null;
		}
		return object;
	}

	/**
	 * This method switches to an alert if there is one.
	 * 
	 * @return True if there is an alert, false if there is no alert.
	 */
	public static boolean isAlertPresent(WebDriver driver) {
		try {
			UserAndPassword up = new UserAndPassword("test", "test2");
			driver.switchTo().alert().authenticateUsing(up);
			return true;
		} catch (NoAlertPresentException e) {
			return false;
		}
	}

	/**
	 * Captures the text of an alert and closes the alert.
	 * 
	 * @return The text of the alert.
	 */
	public static String closeAlertAndGetItsText(WebDriver driver) {
		Alert alert = driver.switchTo().alert();
		String alertText = alert.getText();
		alert.accept();
		return alertText;
	}

} // end TMX library
