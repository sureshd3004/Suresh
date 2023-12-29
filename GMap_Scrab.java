package scrab;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GMap_Scrab{

	public static void main(String[] args) throws IOException, InterruptedException {

		String search = "food+manufacturers+in+MEDAK";


		String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\Arun.xlsx";
		WebDriver driver = new EdgeDriver();
		driver.manage().window().maximize();
		int b = 1;
		int c = 3;
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));
		driver.get("https://www.google.com/search?sca_esv=593729410&rlz=1C1ONGR_enIN1077IN1077&tbs=lf:1,lf_ui:2&tbm=lcl&q=+"+search+"+&rflfq=1&num=100&sa=X&ved=2ahUKEwiFsv276qyDAxWmUGwGHdnhDmcQjGp6BAgWEAE&biw=1466&bih=667&dpr=1.31#rlfi=hd:;si:16691433393574105792,l,Ch1mb29kIG1hbnVmYWN0dXJlcnMga2FyaW1uYWdhckjn_7X-mLSAgAhaLRAAEAEYARgCIh1mb29kIG1hbnVmYWN0dXJlcnMga2FyaW1uYWdhcioECAMQAJIBCnJlc3RhdXJhbnSqAVoKCC9tLzAyd2JtEAEqCCIEZm9vZCgAMh8QASIbI4eTb1XmIN4B5a8MZ9y0s4AsNPyvdqvhIjS3MiEQAiIdZm9vZCBtYW51ZmFjdHVyZXJzIGthcmltbmFnYXI;mv:[[18.4594816,79.163828],[18.3881774,79.1076483]];start:0");

		for (int i=1; i <2000; i++){	

			Thread.sleep(2234);
			WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(19));
			wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id]/div[2]/div/div/a/div/div")));
			List<WebElement> list1 = driver.findElements(By.xpath("//*[@id]/div[2]/div/div/a/div/div"));
			List<WebElement> list2 = driver.findElements(By.xpath("//*[@id]/div[2]/div/div/a/div/div/div[1]/span"));
			List<WebElement> list3= driver.findElements(By.xpath("//*[@id]/div[2]/div/a[1]"));

			if (list1.size() == list2.size() && list2.size() == list3.size()) {

				int size =list1.size();
				WebElement body = driver.findElement(By.tagName("body"));
				body.sendKeys(Keys.END);	
				for (int j = 1; j <size; j++) {
					WebElement elementFromList1 = list2.get(j);
					WebElement elementFromList2 = list1.get(j);
					WebElement elementFromList3 = list3.get(j);
					String zero = elementFromList1.getText();
					String one = elementFromList2.getText().replaceAll(zero,"");
					String two = elementFromList3.getAttribute("href");
					String mobileNumberPattern = "\\b\\d{1,4}[-.\\s]?\\(?" +
							"\\d{1,4}\\)?[-.\\s]?\\d{1,4}[-.\\s]?\\d{1,9}\\b";

					Pattern pattern = Pattern.compile(mobileNumberPattern);
					Matcher matcher = pattern.matcher(one);

					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);
					XSSFSheet sheet = wbook.getSheetAt(0);
					Row row = sheet.createRow(b++);

					org.apache.poi.ss.usermodel.Cell cell1 = row.createCell(1, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell2 = row.createCell(2, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell3 = row.createCell(3, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell4 = row.createCell(4, CellType.STRING);
					cell1.setCellValue(zero);
					cell2.setCellValue(one);
					cell4.setCellValue(two);

					// Find the mobile number in the input string
					if (matcher.find()) {
						String mobileNumber = matcher.group();
						System.out.println("Mobile Number: " + mobileNumber);
						cell3.setCellValue(mobileNumber);
					} 
					//	cell3.setCellValue(mobileNumber);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();				
				} b++;
			}
			WebElement page = driver.findElement(By.xpath("//*[@id='rl_ist0']/div/div[2]/div/table/tbody/tr/td["+c+++"]/a"));
			wait.until(ExpectedConditions.elementToBeClickable(page));
			page.click();
		}
	}			
}