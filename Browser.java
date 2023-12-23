package DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.AssertJUnit;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Browser {
	
	public static WebDriver driver;
	
	
	@BeforeClass
	public static void setUp() {
		
		System.setProperty("webdriver.ie.driver","C:\\Users\\HP\\eclipse-workspace\\Cucumber\\Drivers\\msedgedriver.exe");
		driver =new EdgeDriver();
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);		
			driver.manage().window().maximize();
			
		}
	
	@Test(priority= 0)
public void LoginTest(String user,String pwd,String exp) {
		
		driver.get("https://admin-demo.nopcommerce.com/login");
		WebElement txtdriver =driver.findElement(By.id("Email"));
		txtdriver.clear();
		txtdriver.sendKeys(user);
		
		WebElement txtpwd =driver.findElement(By.id("Password"));
		txtpwd.clear();
		txtpwd.sendKeys(pwd);
		
		driver.findElement(By.xpath("//button[@type='submit']")).click();
		String exp_title="Admin area demo";
		
		String act_title=driver.getTitle();
		
		if(exp.equals("valid"))
		{
			
			if(exp_title.equals(act_title)) {
				
				driver.findElement(By.linkText("Logout")).click();
				AssertJUnit.assertTrue(true);
				
				
			}
			
		}
		else if(exp.equals("invalid"))
		{
			
			if(exp_title.equals(act_title)) {
				
				
				AssertJUnit.assertTrue(false);
			
			}
			
			else
				
			{
				AssertJUnit.assertTrue(true);
				
		}}
		
		
		
	}
	
	
	@Test(priority=1)
	@DataProvider (name="LoginData")
	public void  getData() throws IOException{
		
			File s = new File("C:\\Users\\HP\\eclipse-workspace\\Cucumber\\DataFiles\\data.xlsx");
			FileInputStream ss = new FileInputStream(s);
			Workbook sss = new XSSFWorkbook(ss);
			Sheet t = sss.getSheet("Data");
			int Rows = t.getPhysicalNumberOfRows();
			for (int i = 0; i < Rows; i++) {
				Row row = t.getRow(i);
				int Cells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < Cells; j++) {
					Cell cell = row.getCell(j);
					CellType CT = cell.getCellType();
					if (CT.equals(CellType.STRING)) {
						String s1 = cell.getStringCellValue();
						System.out.println(s1);
					} else if (CT.equals(CellType.NUMERIC)) {
						double numericCellValue = cell.getNumericCellValue();
						int r = (int) numericCellValue;
						System.out.print(r);
					}
					System.out.print("||");
				}
				System.out.println();
				
	}}

			@AfterClass
	public void TearDown() {
		// TODO Auto-generated method stub
				
				driver.close();
		
	}
	

}
