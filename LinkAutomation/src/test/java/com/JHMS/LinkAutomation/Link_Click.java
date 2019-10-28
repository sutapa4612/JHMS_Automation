package com.JHMS.LinkAutomation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Link_Click
{
	WebDriver driver;
	 String spath = System.getProperty("user.dir");
	   
	
	@BeforeMethod
	public void launchurl()
	{
		
		 System.out.println("spath "+spath);
		 System.setProperty("webdriver.chrome.driver", spath+"\\Drivers\\chromedriver.exe");
		driver = new ChromeDriver();
	}
	
	@Test
	public void Link() throws Exception
	{
		 File file=new File(spath+"\\Reources\\links.xlsx");
		String fileName = "links.xlsx";
		String sheetName = "sheet1";
		//Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook book = null;


		//Find the file extension by splitting file name in substring  and getting only extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		//Check condition if the file is xlsx file

		if(fileExtensionName.equals(".xlsx")){

			//If it is xlsx file then create object of XSSFWorkbook class

			book = new XSSFWorkbook(inputStream);

		}

		//Check condition if the file is xls file

		else if(fileExtensionName.equals(".xls")){

			//If it is xls file then create object of HSSFWorkbook class

			book = new HSSFWorkbook(inputStream);

		}

		//Read sheet inside the workbook by its name

		Sheet sheet = book.getSheet(sheetName);

		//Find number of rows in excel file

		int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
        for (int i = 1; i < rowCount+1; i++) {
        	
		Row row = sheet.getRow(i);
			{
				
				org.apache.poi.ss.usermodel.Cell cell=row.getCell(2);
				String cellValue=cell.getStringCellValue();
                System.out.println(cellValue);
                driver.get(cellValue);
                
				//driver.navigate().to(cellValue);
				driver.manage().window().maximize();
				Thread.sleep(5000);
				boolean eleDisplayed;
				try {
					eleDisplayed = driver.findElement(By.xpath("//h1[@id='headingDiv']")).isDisplayed();
					System.out.println("eleDisplayed "+eleDisplayed);
					if(eleDisplayed)
					{
						String sval = driver.findElement(By.xpath("//h1[@id='headingDiv']")).getText();
					    System.out.println("sval "+sval);
						if(sval.equals("Registration Successful"))
						{
							Cell cell1 = sheet.getRow(i).createCell(10);
							cell1.setCellValue(sval);
						}
					}
				} catch (Exception e) {
					// TODO Auto-generated catch block
					//e.printStackTrace();
					System.out.println("sval ");
					Cell cell2 = sheet.getRow(i).createCell(10);
					cell2.setCellValue("Not Success");
				}
				
				FileOutputStream outfile = new FileOutputStream(file);
				book.write(outfile);
				outfile.close();
				  
			}

		}
	}
	
	
	@AfterMethod
	public void closedriver()
	{
		driver.close();
	}
	
}
