import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.io.FileHandler;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;


public class TestCaptcha {
	public static WebDriver driver;
	
	
	
	public static void xcel(int start,int end) throws IOException  
	{  
		
		NumberFormat formatter = new DecimalFormat("#0.00");
		String currentDir = System.getProperty("user.dir");
		String[] excelPath = {currentDir, "\\excel\\data.xlsx"};
		String excelPath1 = String.join(File.separator, excelPath);
		FileInputStream fis = new FileInputStream(excelPath1);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		                        //I have added test data in the cell A1 as "SoftwareTestingMaterial.com"
		                        //Cell A1 = row 0 and column 0. It reads first row as 0 and Column A as 0.
		for(int i=start;i<=end;i++) {
			Row row = sheet.getRow(i);
			
			String enrollment = row.getCell(1).toString();
			String validation = row.getCell(2).toString();
			String entrydate = row.getCell(3).toString();
			String uniqueno = row.getCell(4).toString();
			String officername = row.getCell(5).toString();
			String tlname = row.getCell(6).toString();
			String location = row.getCell(7).toString();
			String capturing = row.getCell(8).toString();
			String checklist = row.getCell(9).toString();
			String causedlist = row.getCell(10).toString();
		
			String enrolllment=formatter.format(Double.parseDouble(enrollment)).substring(0,10);
			String uniquenoo=formatter.format(Double.parseDouble(uniqueno)).substring(0,10);
		
			TestCaptcha tc=new TestCaptcha();
			tc.fillDetails(enrolllment, validation, entrydate, uniquenoo, officername, tlname, location, capturing, checklist, causedlist);
			tc.enterCaptcha();
			
		}
		
	}  
	public void fillDetails(String enroll,String validation,String entrydate,String uni,String officername,String tlname,String location,String capturing,String checklist,String causedlist ) {
		
		
	

		driver.get("https://forms.zohopublic.in/cyberwinlimited/form/ValidationFormQBacked2022JanDKA/formperma/oRosCi9mfffPDQzhds5QsbTdkkcNUlma87YUTPZLI3I");

//	      String l = Keys.chord(Keys.CONTROL,Keys.ENTER);
//	      //open in a new tab
//	      driver.findElement(By.xpath ("//body")).sendKeys(l);
		
		
		
//		driver.navigate().to("https://forms.zohopublic.in/cyberwinlimited/form/ValidationFormQBacked2022JanDKA/formperma/oRosCi9mfffPDQzhds5QsbTdkkcNUlma87YUTPZLI3I");
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[1]/li/div[1]/span[1]/input")).sendKeys(enroll);
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[2]/li/div[1]/div/div/button/em")).click();
		driver.findElement(By.xpath("//div[@class='tempContDiv mSelect']//option[@value='"+validation+"']")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[3]/li/div[1]/div[2]/div/button/em")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[3]/li/div/span[1]/div/input")).sendKeys(entrydate);
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[4]/li/div[1]/div[2]/div/button/em")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[4]/li[1]/div[1]/span[1]/input")).sendKeys(uni);
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[4]/li[2]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown-li']//select//option[@value='"+officername.trim()+"']")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[4]/li[3]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown1-li']//select//option[@value='"+tlname.trim()+"']")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[5]/li/div[1]/div[2]/div/button/em")).click();
		
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[1]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown2-li']//select//option[@value='"+location.trim()+"']")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[2]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown3-li']//select//option[@value='"+capturing.trim()+"']")).click();
		
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[3]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown4-li']//select//option[@value='"+checklist.trim()+"']")).click();
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[4]/div[1]/div[1]/select")).click();
		driver.findElement(By.xpath("//li[@id='Dropdown5-li']//select//option[@value='"+causedlist.trim()+"']")).click();
		
	} 
	public void enterCaptcha() throws IllegalMonitorStateException {
		try {
		WebElement element =driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[5]/div[1]/div/div[2]/img"));
		File src= element.getScreenshotAs(OutputType.FILE);
		String path="C:\\Users\\manu.sharma3\\Selenium_project\\Selenium_project\\captchaimages\\captcha.png";
		FileHandler.copy(src, new File(path));
		Thread.sleep(5000);
		ITesseract image=new Tesseract();
		String code=image.doOCR(new File(path));
		String actualCode="";
		for(int i=0;i<code.length();i++) {
			if(Character.isLetterOrDigit(code.charAt(i))) {
				
				actualCode=actualCode+Character.toString(code.charAt(i));
			}
		}
		System.out.println(actualCode.length());
		if(actualCode.length()==6) {
			
			driver.findElement(By.xpath("//input[@id='verificationcodeTxt']")).sendKeys(actualCode);
			driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[6]/li/div[1]/div[2]/div[2]/button/em")).click();
			driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
			System.out.println("finding error msg");
			while(driver.findElement(By.xpath("/html/body/div[11]/div/div/div")).isDisplayed()) {
				driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
			}
			System.out.println(driver.findElement(By.xpath("//p[@id='error-verificationcode']")).isDisplayed());
			if(driver.findElement(By.xpath("//p[@id='error-verificationcode']")).isDisplayed()) {
				System.out.println("Entering captcha again");
				enterCaptcha();
			}
			}
		else {
			driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[5]/li[5]/div[1]/div/div[1]/div[2]")).click();
			System.out.println("Entering captcha again");
			enterCaptcha();
		}
		}
			catch(Exception e) {
				try {
					if(driver.findElement(By.xpath("//span[@id='splash_msg_text']")).isDisplayed()) {
						String successMessage=driver.findElement(By.xpath("//span[@id='splash_msg_text']")).getText();
						System.out.println(successMessage);
					}
				}
				catch(Exception ee) {
					System.out.println(e);
				}
				System.out.println(e);
			}
			
	}

	public static void main(String[] args) throws InterruptedException, IOException, TesseractException {
		
		TestCaptcha tc=new TestCaptcha();
		String currentDir = System.getProperty("user.dir");
		String[] chromePath = {currentDir, "\\chromedriver.exe"};
		String chromePath1 = String.join(File.separator, chromePath);
		System.setProperty("webdriver.chrome.driver",chromePath1);
		driver=new ChromeDriver();
		
		
		tc.xcel(28141,28150);
		
	}
	
	
}
