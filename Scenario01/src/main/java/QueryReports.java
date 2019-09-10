


import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;

import java.text.ParseException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class QueryReports {

       private static XSSFWorkbook workbook;
       
       
   	public static void main(String [] args) throws SecurityException, IOException, InterruptedException, ParseException, TimeoutException, Exception, WebDriverException {
   	    
   	    FileInputStream fis = new FileInputStream("d:\\Profiles\\aniarora\\Desktop\\Test Automation\\AutomationOutput1.xlsx");
   	    workbook = new XSSFWorkbook(fis);
   	    XSSFSheet sheet = workbook.getSheetAt(0);
   	                                                                   
   	         System.setProperty("webdriver.chrome.driver","d:\\Profiles\\aniarora\\Desktop\\Test Automation\\chromedriver_win32\\chromedriver.exe");                  
   	         WebDriver driver = new ChromeDriver();
   	         
   	 	//	System.setProperty("phantomjs.binary.path","d:\\Profiles\\aniarora\\Desktop\\Test Automation\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe");
   	 	//	WebDriver driver = new PhantomJSDriver();
   	         
   	         driver.get(sheet.getRow(0).getCell(1).getStringCellValue());
   	                 
   	         driver.manage().window().maximize();
   	         driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
   	         
   	         login(sheet.getRow(1).getCell(1).getStringCellValue(), sheet.getRow(2).getCell(1).getStringCellValue(), driver);
   	         System.out.println("Number of Reports in File: " +(sheet.getLastRowNum()-3));
   	      for(int j = 4; j<=sheet.getLastRowNum(); j++)  
   	      {
   	    	String navigation1 = sheet.getRow(j).getCell(1).getStringCellValue();    	
   	    	/*   if(navigation1.equals("End"))
   	        {
   	        	driver.switchTo().defaultContent();
   	        	WebElement home = driver.findElement(By.id("pthdr2home"));
   	        	home.click();
   	        	driver.close();
   	        	System.out.println("End of File");
   	        	break;
   	        } */
   	          System.out.println("Report " +(j-3)+ " Navigation : " +navigation1);
   	          String[] navigation = navigation1.split(" > ");             
   	          String report_name =navigation[navigation.length-1];
   	         try {
   	          mynavigation(navigation, driver);
   	         } catch (Exception navnotfound)
   	         {
   	        	String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
	   			// Call Webdriver to click the screenshot.
	   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

	   			// Save the screenshot.
	   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	        	 System.out.println("Wrong Navigation");
   	        	 sheet.getRow(j).createCell(5).setCellValue("Fail");
   	        	 WebElement home = driver.findElement(By.id("pthdr2home"));
   	        	  	home.click();
   	         }
   	          System.out.println("Report Name: " +report_name);
   	          driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
   	     	 try{
   	     		 driver.switchTo().frame("ptifrmtgtframe");
   	     	 } catch (Exception abcd)
   	     	 {
   	     		 continue;
   	     	 }
   	     	 String case_input;
   	     	if ((report_name.equals("EP Mid Year Review Deputees Rp")) || (report_name.equals("EP Mid Year Review Report")) || 
   	     			(report_name.equals("EP Ful Year Review Report HRBP")) || (report_name.equals("EP Mid Year Review Report HRBP")) || (report_name.equals("PDI Report"))) 
   	     			{
   	     		case_input = "input_year";
   	     			}
   	     	else if ((report_name.equals("Contractors Report")) || (report_name.equals("Marital Status Changes SBS")) || 
   	     			(report_name.equals("Address Change Requests SBS")) || (report_name.equals("Name Change Requests SBS")) || 
   	     			(report_name.equals("SNP Employee List")) || 
   	     			(report_name.equals("Leadership Assess Report HRBP")) || 
   	     			(report_name.equals("Leadership Assessment Defn")) || 
   	     			(report_name.equals("Leadership Assessment Report")) || 
   	     			(report_name.equals("PeopleSoft OTAC mapping"))) 
   	     	{    
   	     		case_input = "no_parameter";
   	     	}
   	     	else if (report_name.equals("Promotion Nomination Report"))
   	     	{
   	     		case_input = "input_year&doc_type";
   	     	}
   	     	else if ((report_name.equals("CDP Deputees Report")) || (report_name.equals("CDP Report")) || 
   	     			(report_name.equals("CDP Report HRBP")))
   	     	{
   	     		case_input = "input_cdp";
   	     	}
   	     	else if(report_name.equals("Sopra Steria Values India"))
   	     	{
   	     		case_input = "input_dropdown_year";
   	     	}
   	     	else if((report_name.equals("Experience Approval Status RPT")) || 
   				(report_name.equals("Qualification Status Report")) || 
   	 			(report_name.equals("Relevant Experience Report")))
   	     	{
   	     		case_input = "input_hrstatus";
   	     	}
   	     	else if(report_name.equals("Employee Address Leavers"))
   	     	{
   	     		case_input = "input_business_unit";
   	     	}
   	     	else if(report_name.equals("Background Verification India"))
   	     	{
   	     		case_input = "input_from_to_doj";
   	     	}     	
   	     	else
   	     	{
   	     		case_input = "input_from_to_years";
   	     	}          
   			switch(case_input) 
   	          {         
   	           case  "input_year" :
   	            { 
   	            try {	
   	        	 input_year(driver, sheet, j);     	 
   	        	 validate(driver, sheet, j);
   	            } catch(Exception er1)
   	            {
   	            	String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	            	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	            	driver.switchTo().defaultContent();
   	            	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	            }
   	        	 break;
   	            }           
   	           case "no_parameter" :
   	            { 
   	            try {	
   	        	 validate(driver, sheet, j); 
   	            } catch(Exception er2)
   	            {
   	            	String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	            	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	            	driver.switchTo().defaultContent();
   	            	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	            }
   	        	 break;
   	            }                
   	           case "input_year&doc_type" :
   	            {
   	           try { 	
   	        	 WebElement p_year = driver.findElement(By.id("InputKeys_YEAR"));
   	        	 String year =  sheet.getRow(j).getCell(2).getStringCellValue();       	 
   	        	 p_year.sendKeys(year);
   	        	 WebElement doc_type = driver.findElement(By.id("InputKeys_EP_REVIEW_TYPE"));
   	        	 doc_type.sendKeys(sheet.getRow(j).getCell(3).getStringCellValue());
   	        	 WebDriverWait wait = new WebDriverWait(driver, 10);                                         
   	        	 WebElement view_result = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("#ICOK")));
   	        	
   	        	 view_result.click();
   	        	 validate(driver, sheet, j);
   	           } catch(Exception er1)
   	           {
   	        	String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
	   			// Call Webdriver to click the screenshot.
	   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

	   			// Save the screenshot.
	   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	           	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	           	driver.switchTo().defaultContent();
   	           	WebElement home = driver.findElement(By.id("pthdr2home"));
   	          	home.click();
   	           }
   	        	 break;
   	            }
   	           case "input_from_to_years" :
   	            { 
   	            	try {
   	        	 if(report_name.equals("Resignation Summary Report"))
   	        	 {
   	        		 Select dropdown_lwd_rd = new Select(driver.findElement(By.id("InputKeys_FLAG7")));
   	        		 dropdown_lwd_rd.selectByVisibleText(sheet.getRow(j).getCell(4).getStringCellValue());
   	        	 }
   	        	 input_from_to_years(driver, sheet, j);            
   	        	 validate(driver, sheet, j);
   	            	} catch(Exception er1)
   	                {
   	            		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
			   			// Call Webdriver to click the screenshot.
			   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

			   			// Save the screenshot.
			   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	                	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	                	driver.switchTo().defaultContent();
   	                	WebElement home = driver.findElement(By.id("pthdr2home"));
   	                  	home.click();
   	                }
   	        	 break;
   	            }
   	           case "input_cdp"  :
   	           {
   	        	   try {
   	        	   input_cdp(driver, sheet, j);        	   
   	        	   validate(driver, sheet, j);
   	        	   } catch(Exception er1)
   	               {
   	        		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	               	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	               	driver.switchTo().defaultContent();
   	               	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	               }
   	        	   break;
   	           }
   	           case "input_dropdown_year" :
   	           {
   	        	   try {        		           	  
   	        	   Select dropdown_year = new Select(driver.findElement(By.id("InputKeys_YEARCD")));
   	        	   dropdown_year.selectByVisibleText(sheet.getRow(j).getCell(2).getStringCellValue());
   	        	   WebElement view_result = driver.findElement(By.id("#ICOK"));
   	        	   view_result.click();
   	        	   validate(driver, sheet, j); 
   	        	   } catch(Exception er1)
   	               {
   	        		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	               	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	               	driver.switchTo().defaultContent();
   	               	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	               }
   	        	   break;
   	           }
   	           case "input_hrstatus" :
   	           {
   	        	   try {
   	        	   input_hrstatus(driver, sheet, j);
   	        	   validate(driver, sheet, j);
   	        	   } catch(Exception er1)
   	               {
   	        		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	               	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	               	driver.switchTo().defaultContent();
   	               	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	               }
   	        	   break;
   	           }
   	           case "input_business_unit" :
   	           {
   	        	   try {
   	        	   WebElement business_unit = driver.findElement(By.id("InputKeys_BUSINESS_UNIT"));
   	        	   business_unit.sendKeys(sheet.getRow(j).getCell(2).getStringCellValue());
   	        	   WebElement view_result = driver.findElement(By.id("#ICOK"));
   	        	   view_result.click();
   	        	   validate(driver, sheet, j);
   	        	   } catch(Exception er1)
   	               {
   	        		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	               	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	               	driver.switchTo().defaultContent();
   	               	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	               }
   	        	   break; 
   	           }
   	           case "input_from_to_doj" :
   	           {
   	        	   try {
   	        	   WebElement from_doj = driver.findElement(By.id("InputKeys_bind2"));
   	        	   from_doj.sendKeys(sheet.getRow(j).getCell(2).getStringCellValue());
   	        	   WebElement to_doj = driver.findElement(By.id("InputKeys_bind3"));
   	        	   to_doj.sendKeys(sheet.getRow(j).getCell(3).getStringCellValue());
   	        	   driver.findElement(By.id("#ICOK")).click();
   	        	   validate(driver, sheet, j);
   	        	   } catch(Exception er1)
   	               {
   	        		String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (j-3) + ".png";
		   			// Call Webdriver to click the screenshot.
		   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

		   			// Save the screenshot.
		   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   	               	sheet.getRow(j).createCell(5).setCellValue("Fail");
   	               	driver.switchTo().defaultContent();
   	               	WebElement home = driver.findElement(By.id("pthdr2home"));
   	              	home.click();
   	               }
   	        	   break;
   	           }
   	         }
   	      }
   	      FileOutputStream fileOut = new FileOutputStream("d:\\Profiles\\aniarora\\Desktop\\Test Automation\\AutomationOutput1.xlsx"); 
   	 	 workbook.write(fileOut); 
   	 		    fileOut.close();
   	    driver.switchTo().defaultContent();
   	  	WebElement home = driver.findElement(By.id("pthdr2home"));
   	  	home.click();
   	  	driver.close();
   	  	System.out.println("End of File");
   	  	
   	    }

   	public static void mynavigation(String[] a,  WebDriver driver)
   	{	
   		
   		driver.switchTo().defaultContent();
   	for(int i = 0; i<a.length; i++)
   	{
   	 WebDriverWait wait_navigation = new WebDriverWait(driver, 10);
   	 WebElement  myElement = wait_navigation.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(a[i])));
   	 myElement.click();
   	}
   	}
   	public static void validate(WebDriver driver, XSSFSheet sheet, int row_number) throws Exception
   	{
   		WebDriverWait wait_validate = new WebDriverWait(driver, 120);

   		WebElement no_values = wait_validate.until(ExpectedConditions.visibilityOfElementLocated(By.className("PSQRYRESULTSTITLE")));
   		 String no_value = no_values.getText();	
   		 int flg1 = 0;
   		 try 
   		 {
   			
   		driver.switchTo().defaultContent();
   		 WebElement ok_btn = driver.findElement(By.cssSelector("input[title = 'Ok (Enter)']"));	
   		 String ErrorScreenshot = System.getProperty("D:\\Profiles\\aniarora\\eclipse-workspace\\Test1") + "ErrorScreenshot" + (row_number-3) + ".png";

   			// Call Webdriver to click the screenshot.
   			File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

   			// Save the screenshot.
   			FileUtils.copyFile(scrFile, new File(ErrorScreenshot));
   		 ok_btn.click();
   		 }
   		 catch (Exception toe1) {	
   			 System.out.println("Message on Screen: " +no_value); 
   			 
   				 flg1 = 1;
   			 }
   		
   		 if (flg1 == 1) 
   		 {
   			 System.out.println("Report Generation Successful");
   			 sheet.getRow(row_number).createCell(5).setCellValue("Pass");
   		 }
   		 else
   		 {
   			 sheet.getRow(row_number).createCell(5).setCellValue("Fail");
   			 System.out.println("Report Generation Unsuccessful"); 
   		 }
   		 WebElement home = driver.findElement(By.id("pthdr2home"));
   		  	home.click();
//   		 FileOutputStream fileOut = new FileOutputStream("d:\\Profiles\\aniarora\\Desktop\\Test Automation\\AutomationOutput.xlsx"); 
//   		 workbook.write(fileOut); 
//   			    fileOut.close();
   	}

   	public static void input_year(WebDriver driver,XSSFSheet sheet, int row_number)
   	{
   	WebElement p_year = driver.findElement(By.id("InputKeys_YEAR"));
   	String year =  sheet.getRow(row_number).getCell(2).getStringCellValue();

   	p_year.sendKeys(year);
   	WebElement view_result = driver.findElement(By.id("#ICOK"));
   	view_result.click();
   	}

   	public static void input_from_to_years(WebDriver driver,XSSFSheet sheet, int row_number)
   	{
   	String date_from = sheet.getRow(row_number).getCell(2).getStringCellValue();                                                              
   	WebElement fromdt = driver.findElement(By.id("InputKeys_bind1"));  
   	fromdt.sendKeys(date_from);     
   	String date_to = sheet.getRow(row_number).getCell(3).getStringCellValue();
   	WebElement todt = driver.findElement(By.id("InputKeys_bind2"));       
   	todt.sendKeys(date_to);
   	WebElement viewresult = driver.findElement(By.id("#ICOK"));
   	viewresult.click();
   	}

   	public static void input_cdp(WebDriver driver,XSSFSheet sheet, int row_number)
   	{
   		WebElement status = driver.findElement(By.id("InputKeys_Z_CP_STATUS"));
   		String status1 =  sheet.getRow(row_number).getCell(2).getStringCellValue();
   		status.sendKeys(status1);
   		WebElement view_result = driver.findElement(By.id("#ICOK"));
   		view_result.click();
   	}

   	public static void login(String id, String pass, WebDriver driver)
   	{
   	    WebElement username = driver.findElement(By.id("userid"));
   	    username.sendKeys(id);
   	    WebElement Password = driver.findElement(By.id("pwd"));
   	    Password.sendKeys(pass);
   	    WebElement loginbutton = driver.findElement(By.name("Submit"));
   	    loginbutton.click();
   	}

   	public static void input_hrstatus(WebDriver driver,XSSFSheet sheet, int row_number)
   	{
   		 Select dropdown_hrstatus = new Select(driver.findElement(By.id("InputKeys_HR_STATUS")));
   		   dropdown_hrstatus.selectByVisibleText(sheet.getRow(row_number).getCell(2).getStringCellValue());
   		   WebElement view_result = driver.findElement(By.id("#ICOK"));
   		   view_result.click();
   	}
   	}