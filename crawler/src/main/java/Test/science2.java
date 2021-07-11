package Test;

import java.io.File;

import java.io.FileOutputStream;

import java.io.IOException;

import java.util.List;

import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;

import org.openqa.selenium.WebElement;

import org.openqa.selenium.chrome.ChromeDriver;

public class science2 {

	public static WebDriver DRIVER;

	public static String Url = "https://www.sciencedirect.com/search?qs=nursing&date=2021&show=100&articleTypes=FLA&offset=0";

	public static Workbook excelWorkbook = null;

	public static String filePath = "C:\\opt\\Email-materials.xlsx";

	public static String sheetName = "emails";

	public static int rowNum = 1;

	public static Workbook wb;

	public static Row row = null;

	public static int j = 0;

	public static void main(String[] args) throws InterruptedException, IOException {

		String filepath = (filePath);

		File file = new File(filepath);

		FileOutputStream fos = new FileOutputStream(file);

		wb = new XSSFWorkbook();

		Sheet sh = wb.createSheet(sheetName);

		row = sh.createRow(0);

		row.createCell(0).setCellValue("SLNO");

		row.createCell(1).setCellValue("F_NAME");

		row.createCell(2).setCellValue("L_NAME");

		row.createCell(3).setCellValue("ADDRESS");

		row.createCell(4).setCellValue("EMAIL ADDRESS");

		row.createCell(0).setCellValue(rowNum);

		System.setProperty("webdriver.chrome.driver", "C:\\opt\\chromedriver.exe");

		DRIVER = new ChromeDriver();

		DRIVER.manage().window().maximize();

		DRIVER.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		DRIVER.get(Url);

		Thread.sleep(1000);

		// DRIVER.findElement(By.xpath("//button[@aria-label=\"close
		// window\"]")).click();

		List<WebElement> links = DRIVER.findElements(By.xpath("//a[contains(@class,'result-list-title-link')]"));

		System.out.println(links.size());

		WebElement page = DRIVER.findElement(By.xpath("//ol[@id=\"srp-pagination\"]/li[1]"));

		String str = page.getText();

		String pagecount = str.substring(str.length() - 2);

		int pagecount1 = Integer.parseInt(pagecount);
		
		try 
		{

		for (int K = 1; K <= pagecount1; K++)

		{
			for (int i = 1; i <= links.size(); i++) // for links
			{

				WebElement links1 = DRIVER

						.findElement(By.xpath("(//a[contains(@class,'result-list-title-link')])[" + i + "]"));

				links1.click();

				Integer emailiconcount = DRIVER

						.findElements(By.xpath("//*[local-name() = 'svg'][@class='icon icon-envelope'][1]")).size();

				System.out.println(emailiconcount);

				Boolean emailiconisPresent1 = emailiconcount > 0;

				System.out.println(emailiconisPresent1);

				if (emailiconisPresent1 == true)

				{

					WebElement emailicon1 = DRIVER

							.findElement(By.xpath("//*[local-name() = 'svg'][@class='icon icon-envelope'][1]"));

					Thread.sleep(800);

					emailicon1.click();

					WebElement fname = DRIVER

							.findElement(
									By.xpath("(//div[@id=\"workspace-author\"]//span[@class='text given-name'])[1]"));

					WebElement lname = DRIVER

							.findElement(By.xpath("(//div[@id=\"workspace-author\"]//span[@class='text surname'])[1]"));

					Row row = sh.createRow(rowNum);

					System.out.println(rowNum);

					row.createCell(0).setCellValue(rowNum);

					row.createCell(1).setCellValue(fname.getText());

					System.out.println(fname.getText());

					row.createCell(2).setCellValue(lname.getText());

					System.out.println(lname.getText());

					WebElement Address = DRIVER

							.findElement(By.xpath("(//div[@id=\"workspace-author\"]//div[@class='affiliation'])[1]"));

					row.createCell(3).setCellValue(Address.getText());

					System.out.println(Address.getText());

					WebElement emailid = DRIVER

							.findElement(By.xpath("(//div[@id=\"workspace-author\"]//div[@class='e-address'])//a"));

					row.createCell(4).setCellValue(emailid.getText());

					System.out.println(emailid.getText());

					Thread.sleep(800);

					DRIVER.findElement(By.xpath("//button[@title=\"Close\"]")).click();
					DRIVER.navigate().back();

					Thread.sleep(800);

					DRIVER.navigate().back();

					DRIVER.navigate().refresh();

				} else {

					row.createCell(1).setCellValue(j);

					row.createCell(2).setCellValue("Email icon is not present");

					System.out.println("Email icon is not present");

					DRIVER.navigate().back();

					DRIVER.navigate().refresh();

				}
				rowNum = rowNum + 1;
			}

			DRIVER.findElement(By.xpath("//a[@data-aa-name=\'srp-next-page\']")).click();
			
			Thread.sleep(800);

			links = DRIVER.findElements(By.xpath("//a[contains(@class,'result-list-title-link')]"));

		}
		} catch( Exception e) {
			System.out.println("Exception: " + e.getStackTrace());
			throw e;
		} finally {
			wb.write(fos);
			fos.close();
		}
	}

}