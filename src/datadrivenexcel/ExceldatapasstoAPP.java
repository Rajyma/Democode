package datadrivenexcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class ExceldatapasstoAPP
{
private WebDriver driver;
private String baseUrl;
private boolean acceptNextAlert = true;
private StringBuffer verificationErrors = new StringBuffer();

@BeforeClass(alwaysRun = true)
public void setUp() throws Exception {
	  
	//create a new work book
			workbook = new HSSFWorkbook();
			//create a new work sheet
			sheet = workbook.createSheet("TestNG Result Summary");
			testresultdata = new LinkedHashMap<String, Object[]>();


			// add test result excel file column header
			// write the header in the first row
			testresultdata.put("1", new Object[] { "Test Step No.", "Action",
					"Expected Output", "Actual Output" });

			  driver = new FirefoxDriver();
			  driver.get("http://52.3.128.41/index.php/site/viewschedule");
			    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			  
}
		 
	@Test(priority=1)
	  public void AddUsers1() throws Exception {
		  
		  
  String company=ExceldatapasstoAPP.getData(0, 0);
  String un=ExceldatapasstoAPP.getData(0, 1);
  String pass=ExceldatapasstoAPP.getData(0, 2);
  System.out.println(company);
  System.out.println(un);
  System.out.println(pass);
  

  
  driver.findElement(By.name("company")).sendKeys(company);
  driver.findElement(By.name("user")).sendKeys(un);
  driver.findElement(By.name("pass")).sendKeys(pass);
  
  System.out.println("login get passed");
  testresultdata.put("2", new Object[] {1d, "navigate to site and login", "site opens and login success","Pass"});
  testresultdata.put("3", new Object[] {2d, "navigate to site and login", "site opens and login success","Pass"});

  testresultdata.put("4", new Object[] {3d, "navigate to site and login", "site opens and login success","Pass"});

  testresultdata.put("5", new Object[] {4d, "navigate to site and login", "site opens and login success","Pass"});

  testresultdata.put("8", new Object[] {8d, "navigate to site and login", "site opens and login success","Pass"});

    }
  

	@Test(priority=2)
	  public void AddUsers2() throws Exception {
		  
		  
String company=ExceldatapasstoAPP.getData(4, 0);
String un=ExceldatapasstoAPP.getData(4, 1);
String pass=ExceldatapasstoAPP.getData(4, 2);
driver.findElement(By.name("company")).clear();
driver.findElement(By.name("company")).sendKeys(company);
driver.findElement(By.name("user")).clear();
driver.findElement(By.name("user")).sendKeys(un);
driver.findElement(By.name("pass")).clear();
driver.findElement(By.name("pass")).sendKeys(pass);
System.out.println(company);
System.out.println(un);
System.out.println(pass);
  
System.out.println("login get passed");
testresultdata.put("10", new Object[] {10d, "navigate to site and login", "site opens and login success","Pass"});
	}
	
	
 public static String getData(int r, int c) throws EncryptedDocumentException, InvalidFormatException, IOException
 {
  FileInputStream FIS=new FileInputStream("G:\\workspace\\sample pro\\src\\testData\\read data1.xlsx");
  Workbook WB=WorkbookFactory.create(FIS);
  String str=WB.getSheet("Sheet1").getRow(r).getCell(c).getStringCellValue();
 
  FileOutputStream Fos=new FileOutputStream("G:\\workspace\\sample pro\\src\\testData\\read data1.xlsx");
	WB.write(Fos);
	Cell cell = null;
	
	 return str;
 
	
 }
 static HSSFWorkbook workbook;
 //define an Excel Work sheet
 HSSFSheet sheet;
  static //define a test result data object
 Map<String, Object[]> testresultdata;

public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
{
	

	
	


}
 

		


	@AfterClass


	public void suiteTearDown() {
		// write excel file and file name is SaveTestNGResultToExcel.xls
		Set<String> keyset = testresultdata.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = testresultdata.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof Date)
					cell.setCellValue((Date) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}
		try {
			FileOutputStream out = new FileOutputStream(new File("SaveTestNGResultToExcel.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Successfully saved Selenium WebDriver TestNG result to Excel File!!!");


		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		// close the browser
		//driver.close();
		//driver.quit();
	
}

}

