package scrab;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

public class SinglePage_MultiElement{
	@Test
    public void mainss() throws InterruptedException, IOException{
        WebDriver driver = new FirefoxDriver();
        driver.get("https://www.synadiet.org/nos-adherents/");
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
        int a =1;
        for (int i = 0; i <=10003; i++) {	
     //   driver.get("https://www.knowde.com/b/technologies-food-ingredients/functional-additives/products/"+i);
        List<WebElement> listurl = driver.findElements(By.xpath("//*[@id=\"list-archive\"]/div/div[2]/div[2]/ul/li/div[2]/div[1]/span")); 
        List<WebElement> listbrand = driver.findElements(By.xpath("//*[@id=\"list-archive\"]/div/div[2]/div[2]/ul/li/div[2]/div[2]/span/font/font")); 
        List<WebElement> listtitle = driver.findElements(By.xpath("//*[@id=\"list-archive\"]/div/div[2]/div[2]/ul/li/div[2]/div[3]/a/font/font"));       
        List<WebElement> loc = driver.findElements(By.xpath("//*[@id=\"list-archive\"]/div/div[2]/div[2]/ul/li/div[1]/span/font/font")); 
    //    List<WebElement> dis = driver.findElements(By.xpath("//*[@id=\"container\"]/div[1]/div[2]/div[2]/table/tbody/tr/td[1]/table/tbody/tr[4]")); 
        int size =listurl.size();
        String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\Cereals.xlsx";     
		FileInputStream fis = new FileInputStream(wpath);
        XSSFWorkbook wbook = new XSSFWorkbook(fis);
        XSSFSheet shee = wbook.getSheet("Sheet2");
        for (int j = 0; j < size; j++) {
        	WebElement elementFromList0 = listtitle.get(j);
            WebElement elementFromList1 = listbrand.get(j);
            WebElement elementFromList2 = listurl.get(j);
            WebElement elementFromList3 = loc.get(j);
    //        WebElement elementFromList4 = dis.get(j);
  //          WebElement elementFromList3 = listchunk.get(j);
            String brand = elementFromList1.getAccessibleName();
            String url = elementFromList2.getAccessibleName();
         //   String head = elementFromList2.getText();
            String title = elementFromList0.getText();  
            String loca = elementFromList3.getText();  
     //       String disc = elementFromList4.getText();  
            System.out.println(url);
            System.out.println(brand);
            Row ro = shee.createRow(i);
            org.apache.poi.ss.usermodel.Cell cell0 = ro.createCell(0, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell1 = ro.createCell(1, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell2 = ro.createCell(2, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell3 = ro.createCell(3, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell4 = ro.createCell(4, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell5 = ro.createCell(5, CellType.STRING);
            org.apache.poi.ss.usermodel.Cell cell6 = ro.createCell(6, CellType.STRING);
            cell0.setCellValue(title);
            cell1.setCellValue(brand);
            cell2.setCellValue(url);
        //    cell3.setCellValue(head);
            cell4.setCellValue(loca);

        } FileOutputStream fos = new FileOutputStream(wpath);
        wbook.write(fos);
        fis.close();
        fos.close();
        wbook.close(); 
        System.out.println("Done");
   driver.quit();
        }
        }}
        	