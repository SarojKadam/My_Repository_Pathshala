package Packg;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.*;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Function;

public class Only_LearnMore_Module {

	static WebDriver driver;

	public static void main(String[] args) throws Exception {
		String learnMoreUrl;
		String eTextUrl;
		File file = new File("F:\\Wynum\\Pathshla\\AdultEducation_189" +
		"\\Adult-Education-In-India.xlsx");//Statistics
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wk = new XSSFWorkbook(fis);
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Shubham\\Downloads\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
		String[] arrayUrl = new String[] {
				    
				"https://epgp.inflibnet.ac.in/view_f.php?category=704"






						
			};
		for (int j = 0; j < arrayUrl.length; j++) {
			
			int totalsheetCount = wk.getNumberOfSheets() + 1;
			XSSFSheet sheet1 = wk.createSheet("Sheet" + totalsheetCount);
			
			driver.get(arrayUrl[j]);
			// fluentWait(By.xpath("//select[@id='module']")).click();
			driver.findElement(By.xpath("//select[@id='module']")).click();
			// taking 1 by 1 Module name using Select class for drop_down
Thread.sleep(1000);
			Select selectModName = new Select(driver.findElement((By.xpath("//select[@id='module']"))));
			// Counting No of Modules from Drop Down
			int noOfModulesName = selectModName.getOptions().size() - 1;
			System.out.println("Total No. Of Module Names: " + noOfModulesName);
			String course, course1;
			XSSFRow row;
			
			if (isElementPresent(By.xpath("/html/body/div/div[3]/div/div/div[1]/h4"))) {
				course = driver.findElement(By.xpath("/html/body/div/div[3]/div/div/div[1]/h4")).getText();
				course1 = course.substring(course.indexOf("Paper:") + 6);
				row = sheet1.createRow(0);
				row.createCell(0).setCellValue(course1);
			}
			row = sheet1.createRow(1);
			row.createCell(0).setCellValue("Module Name");
			row.createCell(1).setCellValue("E-Text Url");
			row.createCell(2).setCellValue("YouTube Url");
			row.createCell(3).setCellValue("Learn More Url");
			//
			for (int i = 1; i <= noOfModulesName; i++) {
				// Taking First Module
				selectModName.selectByIndex(i);
				selectModName.getFirstSelectedOption().click();
				String ModuleName = selectModName.getFirstSelectedOption().getText();
				System.out.println(" Module Name: " + selectModName.getFirstSelectedOption().getText());

				Thread.sleep(2000);
				row = sheet1.createRow(i + 1);
				// XSSFCell cell = row.createCell(i);
				// cell.setCellValue(ModuleName);
				// sheet1.getRow(i).createCell(0).setCellValue(ModuleName);
				row.createCell(0).setCellValue(ModuleName);

				// (2)firstRow secondCol module name
				// sheet1.getRow(i).createCell(i).setCellValue(ModuleName);
//				Thread.sleep(1000);
				// FOR FIRST LABEL T.E.e-text url
				if (isElementPresent(By.id("stop"))) {
					 eTextUrl = driver.findElement(By.xpath("//iframe[@id='video_1']")).getAttribute("src");
					// System.out.println("eTextUrl url: " + eTextUrl);
					// sheet1.getRow(i).createCell(1).setCellValue(eTextUrl);
					// sheet1.getRow(rownum)cell.setCellValue(eTextUrl);
				//Thread.sleep(1000);	
					row.createCell(1).setCellValue(eTextUrl);
					System.out.println("eTextUrl" + eTextUrl);

				} else {
					row.createCell(1).setCellValue("E-TextUrl is Absent");
					//sheet1.getRow(i).createCell(1).setCellValue("E-TextUrl is Absent");
				}
				// FOR SECOND T.E. SELF LEARNING T.E. youtube url
				if (isElementPresent(By.id("video_2"))) {
					String youtubeLink = driver.findElement(By.xpath("//iframe[@id='youtube_player']"))
							.getAttribute("src");
					// sheet1.getRow(i).createCell(2).setCellValue(youtubeLink);
					row.createCell(2).setCellValue(youtubeLink);
					System.out.println("YouTube Url is:" + youtubeLink);
				} else {
					row.createCell(2).setCellValue("You Tube Url is Absent");
					// sheet1.getRow(i).createCell(2).setCellValue("You Tube Url is Absent");
					// System.out.println("Video is absent");
				}

				// FOR THIRD T.E. LEARN MORE LABEL
				if (isElementPresent(By.id("stop2"))) {

					if (isElementPresent(By.id("stop")) && isElementPresent(By.id("video_2"))
							&& isElementPresent(By.id("stop1"))) {
						learnMoreUrl = driver
								.findElement(By.xpath("//*[@id='module_type_content']/div[2]/div[4]/iframe"))
								.getAttribute("src");

						row.createCell(3).setCellValue(learnMoreUrl);
						System.out.println("LearneMore Url=" + learnMoreUrl);
					}
					// sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl);

					else if (isElementPresent(By.id("stop")) && isElementPresent(By.id("video_2"))) {
						learnMoreUrl = driver
								.findElement(By.xpath("//*[@id='module_type_content']/div[2]/div[3]/iframe"))
								.getAttribute("src");

						row.createCell(3).setCellValue(learnMoreUrl);
						System.out.println("LearneMore Url=" + learnMoreUrl);
						// System.out.println("LearneMore Url=" + learnMoreUrl);

					} else if (isElementPresent(By.id("stop")) && isElementPresent(By.id("stop1"))) {
						learnMoreUrl = driver
								.findElement(By.xpath("//*[@id='module_type_content']/div[2]/div[3]/iframe"))
								.getAttribute("src");

						row.createCell(3).setCellValue(learnMoreUrl);
						System.out.println("LearneMore Url=" + learnMoreUrl);
						// sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl);
						// System.out.println("LearneMore Url=" + learnMoreUrl);

					} else if (isElementPresent(By.id("video_2")) && isElementPresent(By.id("stop1"))) {
						learnMoreUrl = driver
								.findElement(By.xpath("//*[@id='module_type_content']/div[2]/div[3]/iframe"))
								.getAttribute("src");

						row.createCell(3).setCellValue(learnMoreUrl);
						// sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl);
						System.out.println("LearneMore Url=" + learnMoreUrl);

					} 
					else if (isElementPresent(By.id("stop")) &&
							  isElementPresent(By.id("stop2"))) {
							  learnMoreUrl = driver.findElement(By.xpath(
							  "//*[@id='module_type_content']/div[2]/div[2]/iframe")) .getAttribute("src");
							  row.createCell(3).setCellValue(learnMoreUrl); //
							  sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl); }
					else if (isElementPresent(By.id("video_2")) &&
							  isElementPresent(By.id("stop2"))) {
							  learnMoreUrl = driver.findElement(By.xpath(
							  "//*[@id='module_type_content']/div[2]/div[2]/iframe")) .getAttribute("src");
							  row.createCell(3).setCellValue(learnMoreUrl); //
							  sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl); }
					else if (isElementPresent(By.id("stop1")) &&
							  isElementPresent(By.id("stop2"))) {
							  learnMoreUrl = driver.findElement(By.xpath(
							  "//*[@id='module_type_content']/div[2]/div[2]/iframe")) .getAttribute("src");
							  row.createCell(3).setCellValue(learnMoreUrl); //
							  sheet1.getRow(i).createCell(3).setCellValue(learnMoreUrl); }
					else if (isElementPresent(By.id("stop1"))) {
						learnMoreUrl = driver
								.findElement(By.xpath("//*[@id='module_type_content']/div[2]/div[1]/iframe"))
								.getAttribute("src");
						row.createCell(3).setCellValue(learnMoreUrl);
						System.out.println("LearneMore Url=" + learnMoreUrl);
					}
					
					 
					 
				} else {
					row.createCell(3).setCellValue("LearnMore Url is Absent");
					// sheet1.getRow(i).createCell(3).setCellValue("LearnMore Url is Absent");
				}
				FileOutputStream fos = new FileOutputStream(file);
				wk.write(fos);

			} // For loop
		} // first / url for loop

		/*
		 * FileOutputStream fos=new FileOutputStream(file); wk.write(fos); wk.close();
		 */
		// System.out.println( );
		driver.close();
		wk.close();
	}// main method

	private static boolean isElementPresent(By id) throws InterruptedException {
		try {
			//WebElement element = (new WebDriverWait(driver, 10))
			//		   .until(ExpectedConditions.elementToBeClickable(id));
			
			if(id.toString()=="stop")
			{
				Thread.sleep(2000);
			}
			driver.findElement(id);
			return true;
		} catch (org.openqa.selenium.NoSuchElementException e) {
			return false;
		}
	}// isElementPresent

	public static WebElement fluentWait(final By locator) {

		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(Duration.ofSeconds(30))
				.pollingEvery(Duration.ofSeconds(2)).ignoring(NoSuchElementException.class)
				.ignoring(NoSuchMethodError.class);

		WebElement foo = wait.until(new Function<WebDriver, WebElement>() {
			public WebElement apply(WebDriver driver) {
				return driver.findElement(locator);
			}
		});
		return foo;

	}
}
