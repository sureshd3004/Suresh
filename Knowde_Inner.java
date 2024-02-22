package year2022;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

public class Knowde_Inner{
	@Test
	public void mainss() throws InterruptedException, IOException{
		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));   

		for (int i =2056; i <=100003; i++) {	
			XSSFWorkbook rbook = new XSSFWorkbook("C:\\Users\\sures\\OneDrive\\Desktop\\NEW1.xlsx");
			XSSFSheet sheet = rbook.getSheet("DATA");
			XSSFRow li = sheet.getRow(i);
			XSSFCell cel2 = li.getCell(2);
			String value = cel2.getStringCellValue();
			driver.get(value);
			rbook.close();            
			System.out.println(i);
			try {
				WebElement elementFromList0 = driver.findElement(By.xpath("//*[@id=\"description\"]/div/div[1]/knowde-product-summary")); 
				if (elementFromList0 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String title = elementFromList0.getText();
					Row ro = shee.createRow(i);
					org.apache.poi.ss.usermodel.Cell cell0 = ro.createCell(4, CellType.STRING);
					cell0.setCellValue(title);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}
			} catch (NoSuchElementException e) {
				System.out.println("No Element-o");
			}try {
				WebElement elementFromList1 = driver.findElement(By.xpath("//*[@id=\"description\"]/div/div[1]/knowde-product-page-header/div/div/div[2]"));
				if (elementFromList1 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String brand = elementFromList1.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell1 = ro.createCell(3, CellType.STRING);
					cell1.setCellValue(brand);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}
			} catch (NoSuchElementException e) {
				System.out.println("No Element-1");
			}try {
				WebElement elementFromList2 = driver.findElement(By.xpath("//*[@id='identification-&-functionality']/section"));
				if (elementFromList2 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String url =   elementFromList2.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell2 = ro.createCell(5, CellType.STRING);
					cell2.setCellValue(url);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}
			} catch (NoSuchElementException e) {
				System.out.println("No Element-2");
			}try {
				WebElement elementFromList3 = driver.findElement(By.xpath("//*[@id='features-&-benefits']/section"));
				if (elementFromList3 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String whole = elementFromList3.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell3 = ro.createCell(6, CellType.STRING);
					cell3.setCellValue(whole);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}
			} catch (NoSuchElementException e) {
				System.out.println("No Element-3");
			}try {
				WebElement elementFromList4 = driver.findElement(By.xpath("//*[@id='applications-&-uses']/section")); 
				if (elementFromList4 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String titl =  elementFromList4.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell4 = ro.createCell(7, CellType.STRING);
					cell4.setCellValue(titl);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}   
			} catch (NoSuchElementException e) {
				System.out.println("No Element-4");
			} 
			try {
				WebElement elementFromList4 = driver.findElement(By.xpath("//*[@id='packaging-&-availability']")); 
				if (elementFromList4 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String titl =  elementFromList4.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell4 = ro.createCell(8, CellType.STRING);
					cell4.setCellValue(titl);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}   
			} catch (NoSuchElementException e) {
				System.out.println("No Element-5");
			} 
			try {
				WebElement elementFromList4 = driver.findElement(By.xpath("//*[@id=\"storage-&-handling\"]")); 
				if (elementFromList4 != null) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					String titl =  elementFromList4.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell4 = ro.createCell(9, CellType.STRING);
					cell4.setCellValue(titl);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();

				}   
			} catch (NoSuchElementException e) {
				System.out.println("No Element-6");
			} 

			int yen=10;
			try {
				List<WebElement> elementFromList5 = driver.findElements(By.xpath("//*[@id=\"properties\"]/section/div[2]/div/knowde-content-block")); 
				for (int j = 1; j < elementFromList5.size(); j++) {
					String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\FSSI WRITE.xlsx";     
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet shee = wbook.getSheet("Sheet1");
					WebElement List5 = driver.findElement(By.xpath("//*[@id=\"properties\"]/section/div[2]/div["+j+"]/knowde-content-block")); 
					String loop = List5.getText();
					Row ro = shee.getRow(i);
					org.apache.poi.ss.usermodel.Cell cell9 = ro.createCell(yen++, CellType.STRING);
					cell9.setCellValue(loop);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();
					wbook.close(); 
					
				}   
			} catch (NoSuchElementException e) {
				System.out.println("No Element-7");
			}     
		}     
	}
}
