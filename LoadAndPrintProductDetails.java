package scrab;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class LoadAndPrintProductDetails {
	
	@Test
    public void man() throws InterruptedException, IOException {
        WebDriver driver = new FirefoxDriver();
        driver.get("https://www.ingredientsnetwork.com/live/search/searchresults46v2.jsp?name=&site=47&SugType_val=&RecordId_val=&searchtype=products");
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(18));
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(19));
     
        try {
            By loadMoreButtonLocator = By.xpath("/html/body/div[2]/div[2]/div[1]/div[2]/div[3]");
            
        //    System.out.println(driver.findElement(loadMoreButtonLocator).getLocation());
            for (int i = 0; i <10000; i++) {
            	//Thread.sleep(5000);
                WebElement loadMoreButton = driver.findElement(loadMoreButtonLocator);            
                wait.until(ExpectedConditions.elementToBeClickable(loadMoreButton));
                JavascriptExecutor executor = (JavascriptExecutor) driver;
        		executor.executeScript("533", "7654");
            System.out.println(i);
                loadMoreButton.click();
        }
        } catch (Exception e) { System.out.println("catch");  }
        finally {
            Thread.sleep(2000);          
            List<WebElement> productNames = driver.findElements(By.xpath("//*[@id]/a/div/h3"));
            List<WebElement> productchunk = driver.findElements(By.xpath("//*[@id]/a/div"));
            List<WebElement> brandNames = driver.findElements(By.xpath("//*[@id]/a/div/p[1]/span"));   
            List<WebElement> brandNa = driver.findElements(By.xpath("/html/body/div[2]/div[2]/div[1]/div[2]/div[2]/div/a"));   
            String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\404.xlsx";     
    		FileInputStream fis = new FileInputStream(wpath);
            XSSFWorkbook wbook = new XSSFWorkbook(fis);
            XSSFSheet shee = wbook.getSheet("Sheet6");
            int a = 1;
            for (int i = 0; i < brandNa.size(); i++) {
                String productName = productNames.get(i).getText();
                String productchunck = productchunk.get(i).getText();
                String brandName = brandNames.get(i).getText();
                String url = brandNa.get(i).getAttribute("href");
                Row ro = shee.createRow(a++);
                org.apache.poi.ss.usermodel.Cell cell0 = ro.createCell(0, CellType.STRING);
                org.apache.poi.ss.usermodel.Cell cell1 = ro.createCell(1, CellType.STRING);
                org.apache.poi.ss.usermodel.Cell cell2 = ro.createCell(2, CellType.STRING);  
                org.apache.poi.ss.usermodel.Cell cell3 = ro.createCell(3, CellType.STRING);  
                cell0.setCellValue(productName);
                cell1.setCellValue(brandName);
                cell2.setCellValue(productchunck);  
                cell3.setCellValue(url);
            
            }   FileOutputStream fos = new FileOutputStream(wpath);
            wbook.write(fos);
            fis.close();
            fos.close();
            wbook.close(); 
        }}  
    @Test
    private static boolean isElementPresent(WebDriver driver, By by) {
        try {
            driver.findElement(by);
            return true;
        } catch (org.openqa.selenium.NoSuchElementException e) {
            return false;
        }}}
