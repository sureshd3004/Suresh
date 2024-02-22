package year2022;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class E2E_EU_Variable_v2 {

	static String MailID = "suresh.d@12taste.com";
	static String Password ="suresh2508";
	static String filepath="C:\\Users\\sures\\OneDrive\\Desktop\\404.xlsx";
	static String sheetname="Sheet8";
	public static WebDriver driver;
	public static Select dropdown;

	public static void main(String[] args) throws IOException, InterruptedException {

		driver = new EdgeDriver();
		driver.get("https://www.12taste.com/my-account/");
		driver.manage().window().maximize();
		// ((JavascriptExecutor) driver).executeScript("document.body.style.zoom='80%'");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
		driver.findElement(By.id("username")).sendKeys("suresh.d@12taste.com");
		driver.findElement(By.id("password")).sendKeys("suresh2508");
		driver.findElement(By.xpath("//*[@id=\"customer_login\"]/div[1]/form/p[3]/button")).click();
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(19));

		for (int i=1; i <=3000; i++){
			FileInputStream fis = new FileInputStream(filepath);
			XSSFWorkbook rbook = new XSSFWorkbook(filepath);
			XSSFSheet sheet = rbook.getSheet(sheetname);
			XSSFRow li = sheet.getRow(i);
			
			XSSFCell cell27 = li.createCell(27);           

			XSSFCell cell1 = li.getCell(1);          XSSFCell cell0 = li.getCell(0);
			XSSFCell cell2 = li.getCell(2);          XSSFCell cell3 = li.getCell(3);
			XSSFCell cell4 = li.getCell(4);          XSSFCell cell5 = li.getCell(5);			
			XSSFCell cell6 = li.getCell(6);		     XSSFCell cell7 = li.getCell(7);    	XSSFCell cell8 = li.getCell(8);
			XSSFCell cell9 = li.getCell(9);		     XSSFCell cell10= li.getCell(10);       XSSFCell cell11= li.getCell(11);
			XSSFCell cell12= li.getCell(12);         XSSFCell cell13= li.getCell(13);		XSSFCell cell14= li.getCell(14);			
			XSSFCell cell15= li.getCell(15);         XSSFCell cell16 = li.getCell(16);      XSSFCell cell17 = li.getCell(17); 
			XSSFCell cell18 = li.getCell(18);        XSSFCell cell19 = li.getCell(19);      XSSFCell cell20 = li.getCell(20);
			XSSFCell cell21 = li.getCell(21);        XSSFCell cell22 = li.getCell(22);		XSSFCell cell23 = li.getCell(23);
			XSSFCell cell24 = li.getCell(24);        XSSFCell cell25 = li.getCell(25);      XSSFCell cell26 = li.getCell(26);

			String Name                 = cell0.getStringCellValue();         
			String FlavourProfile       = cell8.getStringCellValue();               
			String shortdiscription     = cell1.getStringCellValue();        String ShelfLife = cell9.getRawValue();
			String longdiscription      = cell2.getStringCellValue();        String StorageConditions = cell10.getStringCellValue();
			String applicationarea      = cell3.getStringCellValue();        String Technology = cell11.getStringCellValue();
			XSSFRichTextString HSNCod   = cell4.getRichStringCellValue();    String HSNCode=HSNCod.toString();       String DGClassification = cell12.getStringCellValue();
			String CountryofOrigin      = cell5.getStringCellValue();        String DietarySuitability = cell13.getStringCellValue();
			String LabelsandMarks       = cell6.getStringCellValue();        String Appearance = cell14.getStringCellValue();
			String FlavourDescriptors   = cell7.getStringCellValue();        String shipment = cell15.getStringCellValue();
			String tax = cell16.getStringCellValue();                        String vendor = cell17.getStringCellValue();
			String weight = cell18.getRawValue();                            String minimumQuantity = cell21.getRawValue();
			String image =  cell20.getStringCellValue();	                 String PDS_URL = cell19.getStringCellValue();
			String maximumQuantity = cell22.getRawValue();                   String Commission=cell23.getRawValue();
			String ReggularPrice=cell24.getRawValue();                       String SalePrice=cell25.getRawValue();
			String DelivaryNote = "2 days to delivary";
			String Packaging = cell26.getStringCellValue();           //       String productname = cell27.getStringCellValue();

			driver.get("https://www.12taste.com/wp-admin/post-new.php?post_type=product");
		
			Thread.sleep(2000);
			//   wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='variable_product_options_inner']/div[3]/button[2]")));
			driver.findElement(By.xpath("//span[text()='Inventory']")).click();
			WebElement sku = driver.findElement(By.xpath("//*[@id=\"_sku\"]"));
			String id = sku.getAttribute("value");
			
			cell27.setCellValue(id);
			FileOutputStream fos = new FileOutputStream(filepath);
			rbook.write(fos);
			fis.close();
			fos.close();
			rbook.close();	
			driver.get("https://www.12taste.com/wp-admin/post.php?post="+id+"&action=edit");
			driver.findElement(By.xpath("//*[@id='title']")).sendKeys(Name);
			WebElement elem = driver.findElement(By.xpath("//*[@id='product-type']"));
			Select sel = new Select(elem);
			sel.selectByVisibleText("Variable product");
			try {	
				driver.findElement(By.xpath("//span[text()='General']")).click();
				Thread.sleep(1500);
				WebElement dropdownelement = driver.findElement(By.xpath("//*[@id='_tax_class']"));
				WebElement dropdowne = driver.findElement(By.xpath("//*[@id='_tax_status']"));
				dropdown = new Select(dropdownelement);
				Select dde = new Select(dropdowne);
				dropdown.selectByVisibleText(tax);
				dde.selectByVisibleText("Taxable");
				driver.findElement(By.xpath("//*[@id='wcpv_customize_product_vendor_settings']")).click();
				driver.findElement(By.xpath("//*[@id='wcpv-keep-tax']")).click();
			} catch (Exception e) {
				System.out.println("Tax");
				e.printStackTrace();
			}
			try{
				driver.findElement(By.xpath("//span[text()='Shipping']")).click();
				//	driver.findElement(By.xpath("//*[@id='woocommerce-product-data']/div[2]/div/ul/li[3]/a")).click();
				driver.findElement(By.xpath("//*[@id='_weight']")).sendKeys(weight);
				WebElement dropdownelement = driver.findElement(By.xpath("//*[@id='product_shipping_class']"));
				dropdown = new Select(dropdownelement);
				dropdown.selectByVisibleText(shipment);
			} catch (Exception e) {
				System.err.println("shipping");
				e.printStackTrace();
			}
			try { 	
				WebElement att = driver.findElement(By.xpath("//span[text()='Attributes']"));
				att.click();	
				int packaging =1 ;
				Thread.sleep(4000);
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']"))).click();
				//	WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));		
				//	s.click();
				WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
				sd.sendKeys("packa");
				sd.click();
				driver.findElement(By.xpath("//li[text()='Packaging']")).click();
				WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_packaging.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));     
				String[] wordsArray = Packaging.split(",");
				for (String String_word : wordsArray){
					Thread.sleep(2000);
					wait.until(ExpectedConditions.elementToBeClickable(a)).click();
					a.sendKeys(String_word);
					driver.findElement(By.xpath("//li[@class='select2-results__option select2-results__option--highlighted']")).click();				    
				}//	driver.findElement(By.xpath("//input[@name='attribute_visibility[0]']")).click();			
				driver.findElement(By.xpath("//*[@id='product_attributes']/div[2]/div[1]/h3/div[1]")).click();
				Thread.sleep(1000);
				driver.findElement(By.xpath("//*[@id='product_attributes']/div[3]/button")).click();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Packaging");
				e.printStackTrace();
			}
		

			int Attribute =2 ;
			try {				
				
				Thread.sleep(3000);
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
				WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));		
				s.click();
				WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
				sd.sendKeys("application");
				sd.click();
				driver.findElement(By.xpath("//li[text()='Application Area']")).click();
				WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_application-area.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));     
				String[] wordsArray = applicationarea.split(",");
				for (String String_word : wordsArray){
					a.sendKeys(String_word);
					Thread.sleep(4000);
					driver.findElement(By.xpath("//li[@class='select2-results__option select2-results__option--highlighted']")).click();
				}driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[2]/div/table/tbody/tr[3]/td/div/label/input")).click();
				driver.findElement(By.xpath("//*[@id='product_attributes']/div[2]/div[2]/h3/div[1]")).click();						
			} catch (Exception e) {
				System.out.println("Application");
				e.printStackTrace();
			}
			try {		
				int DG =3 ;
				if (DGClassification.equalsIgnoreCase("null")) {}				   
				else {
					//	WebElement att = driver.findElement(By.xpath("//span[text()='Attributes']"));
					//	att.click();
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("DG");
					sd.click();
					driver.findElement(By.xpath("//li[text()='DG Classification']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_dg-classification.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));     
					a.sendKeys(DGClassification);				
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[3]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[3]/h3/div[1]")).click();
				}} catch (Exception e) {
					System.out.println("DG");
					e.printStackTrace();
				}
			//Thread.sleep(500);
			try {

				if (DietarySuitability.equalsIgnoreCase("null")) {}				   
				else {
					int Dietary =4 ;
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Dietary Suitability");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Dietary Suitability']")).click();			
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_dietary-certifications.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));     
					a.sendKeys(DietarySuitability);				
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[4]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[4]/h3/div[1]")).click();
				} }catch (Exception e) {
					System.out.println("Dietary Suitability");
					e.printStackTrace();
				}//Thread.sleep(500);
			try {
				if (Appearance.equalsIgnoreCase("null")) {}				   
				else {
					int appearance =5 ;
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Appear");
					sd.click();
					driver.findElement(By.xpath("//*[@class='select2-results__option select2-results__option--highlighted']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_appearance.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));     
					a.sendKeys(Appearance);				
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[5]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[5]/h3/div[1]")).click();
				}} catch (Exception e) {
					System.out.println("Appearance");
					e.printStackTrace();
				}//Thread.sleep(500);
			try {
				if (CountryofOrigin.equalsIgnoreCase("null")) {}				   
				else {
					int Country =6 ;
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Country of Origin");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Country of Origin']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_country-of-origin.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));     
					a.sendKeys(CountryofOrigin);				
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[6]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[6]/h3/div[1]")).click();
				}} catch (Exception e) {
					System.out.println("Country of Origin");
					e.printStackTrace();
				}//Thread.sleep(500);
			try {
				if (ShelfLife.equalsIgnoreCase("null")) {}				   
				else {
					int shelfLife =7 ;
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Shelf Life");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Shelf Life']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_shelf-life.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));     
					a.sendKeys(ShelfLife);				
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[7]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[7]/h3/div[1]")).click();
				}} catch (Exception e) {
					System.out.println("Shelf Life");
					e.printStackTrace();
				}
			try {
				int hsn =8 ;
				if (HSNCode.equalsIgnoreCase("null")) {}				   
				else {
					//Thread.sleep(500);
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("HSN");
					sd.click();
					driver.findElement(By.xpath("//li[text()='HSN Code']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_hsn-code.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));         
					a.sendKeys(HSNCode);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[8]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[8]/h3/div[1]")).click();
				} }catch (Exception e) {
					System.out.println("HSN Code");
					e.printStackTrace();
				}
			try {
				if (LabelsandMarks.equalsIgnoreCase("null")) {}				   
				else {
					//	driver.findElement(By.xpath("//span[text()='Attributes']")).click();
					//	Thread.sleep(500);
					int labels =9 ;
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("labels");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Labels and Marks']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_labels-and-marks.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));         
					a.sendKeys(LabelsandMarks);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[9]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[9]/h3/div[1]")).click();
				} }catch (Exception e) {
					System.out.println("Labels and Marks");
					e.printStackTrace();
				}
			try {
				if (FlavourProfile.equalsIgnoreCase("null")) {}				   
				else {
					//	driver.findElement(By.xpath("//span[text()='Attributes']")).click();
					//	Thread.sleep(500);
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Flavour Profile");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Flavour Profile']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_flavour-profile.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li > input"));         
					a.sendKeys(FlavourProfile);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[10]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id='product_attributes']/div[2]/div[10]/h3/div[1]")).click();
				} }catch (Exception e) {
					System.out.println("Flavour Profile");
					e.printStackTrace();
				}
			try {
				if (Technology.equalsIgnoreCase("null")) {}				   
				else {
					//	Thread.sleep(500);
					JavascriptExecutor executor1 = (JavascriptExecutor) driver;			
					executor1.executeScript("arguments[0].scrollIntoView();", elem);
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Technology");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Technology']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_technology.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));         
					a.sendKeys(Technology);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[11]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[11]/h3/div[1]")).click();		
				}} catch (Exception e) {
					System.out.println("Technology");
					e.printStackTrace();
				}
			try {
				if (StorageConditions.equalsIgnoreCase("null")) {}				   
				else {
					JavascriptExecutor executor1 = (JavascriptExecutor) driver;			
					executor1.executeScript("arguments[0].scrollIntoView();", elem);
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Storage Conditions");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Storage Conditions']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_storage-conditions.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));         
					a.sendKeys(StorageConditions);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[12]/div/table/tbody/tr[3]/td/div/label/input")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[12]/h3/div[1]")).click();
				}} catch (Exception e) {
					System.out.println("Storage Conditions");
					e.printStackTrace();
				}
			try {
				if (FlavourDescriptors.equalsIgnoreCase("null")) {}				   
				else {
					JavascriptExecutor executor1 = (JavascriptExecutor) driver;			
					executor1.executeScript("arguments[0].scrollIntoView();", elem);
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='select2-selection__placeholder']")));
					WebElement s = driver.findElement(By.xpath("//span[@class='select2-selection__placeholder']"));
					s.click();
					WebElement sd = driver.findElement(By.xpath("/html/body/span/span/span[1]/input"));
					sd.sendKeys("Flavour Descriptors");
					sd.click();
					driver.findElement(By.xpath("//li[text()='Flavour Descriptors']")).click();
					WebElement a = driver.findElement(By.cssSelector("#product_attributes > div.product_attributes.wc-metaboxes.ui-sortable > div.woocommerce_attribute.wc-metabox.postbox.taxonomy.pa_flavour-descriptors.open > div > table > tbody > tr:nth-child(1) > td:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input"));         
					a.sendKeys(FlavourDescriptors);
					driver.findElement(By.xpath("//*[@class=\"select2-results__option select2-results__option--highlighted\"]")).click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[13]/div/table/tbody/tr[3]/td/div/label/input")).click();
					WebElement element = driver.findElement(By.xpath("//*[@id='product_attributes']/div[3]/button"));
					element.click();
					driver.findElement(By.xpath("//*[@id=\"product_attributes\"]/div[2]/div[13]/h3/div[1]")).click();
					Thread.sleep(1234);
				}} catch (Exception e) {
					System.out.println("Flavour Descriptors");
					e.printStackTrace();
				}
			 ((JavascriptExecutor) driver).executeScript("document.body.style.zoom='50%'");
			try {			
				WebElement ld = driver.findElement(By.xpath("//*[@id='content']"));	
				ld.clear();
				ld.sendKeys(longdiscription);
			} catch (Exception e) {
				System.err.println("long Discription");
				e.printStackTrace();
			}
			try {	
				WebElement body = driver.findElement(By.tagName("body"));
				body.sendKeys(Keys.END);
				WebElement sd = driver.findElement(By.xpath("//*[@id='excerpt']"));		
				sd.clear();
				sd.sendKeys(shortdiscription);
			} catch (Exception e) {
				System.err.println("short Discription");
				e.printStackTrace();
			}	
			 ((JavascriptExecutor) driver).executeScript("document.body.style.zoom='100%'");
			try{			
				driver.findElement(By.xpath("//*[@id=\"wcpv-product-vendor-terms-select\"]")).sendKeys(vendor);	
			} catch (Exception e) {
				System.err.println("vendor");
				e.printStackTrace();
			}
			try {			
				driver.findElement(By.xpath("//span[text()='Product Documents']")).click();
				WebElement s = driver.findElement(By.xpath("//button[text()='New Section']"));
				s.click();
				WebElement ss = driver.findElement(By.xpath("//button[text()='Add Document']"));
				ss.click();
				WebElement label = driver.findElement(By.xpath("//*[@id=\"wc_product_document_label_0_0\"]"));
				label.sendKeys("Safety data sheet");
				WebElement filepath = driver.findElement(By.xpath("//*[@id=\"wc_product_document_file_location_0_0\"]"));
				filepath.sendKeys(PDS_URL);
				WebElement name = driver.findElement(By.xpath("//*[@id=\"product_documents_section_name_0\"]"));
				name.sendKeys("Product Document");
				driver.findElement(By.xpath("//*[@id=\"_wc_product_documents_display\"]")).click();	
				Thread.sleep(1234);
			}   catch (Exception e) {
				System.err.println("PDS");
				e.printStackTrace();
			}
			/*	try {
			List<WebElement> cat = driver.findElements(By.xpath("//*[@id]/label"));
			for (WebElement webElement : cat) {
				String ids = webElement.getAttribute("id");
				String name = webElement.getText();
				System.out.println(name);
				System.out.println(ids);
			}
		} catch (Exception e) {
			System.out.println("cat");
		}            */
		/*	try {		
				driver.get("https://www.12taste.com/wp-admin/media-upload.php?post_id="+id+"&type=image&TB_iframe=1&width=753&height=546");
				WebElement fileInput1 = driver.findElement(By.xpath("//*[@id=\"async-upload\"]"));				                                                                                                         //Set The File Path
				fileInput1.sendKeys(image);
				driver.findElement(By.xpath("//*[@id=\"html-upload\"]")).click();
				WebElement  done = driver.findElement(By.xpath("//*[@class='wp-post-thumbnail']"));
				done.click();	
				driver.get("https://www.12taste.com/wp-admin/post.php?post="+id+"&action=edit");
			} catch (Exception e) {
				System.err.println("Image");
				e.printStackTrace();
			}    */
			try {	
				//	int packaging =3 ;
				WebElement att = driver.findElement(By.xpath("//span[text()='Variations']"));
				att.click();
				Thread.sleep(1000);
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='variable_product_options_inner']/div[2]/button[2]")));
				Thread.sleep(1000);
				WebElement addManually = driver.findElement(By.xpath("//*[@id='variable_product_options_inner']/div[2]/button[2]"));
				String[] wordsArray = Packaging.split(",");
				String[] ReggularPric = ReggularPrice.split(",");
				String[] SalePric= SalePrice.split(",");
				String[] Commissio = Commission.split(",");
				String[] Weight = weight.split(",");

				String one = wordsArray[2];
				int y =0;
				for (String String_word : wordsArray) {					
					for (int j = 1; j <2; j++) {
						wait.until(ExpectedConditions.elementToBeClickable(addManually));
						Thread.sleep(1000);
						addManually.click();	                                      
						Thread.sleep(3000);
						//	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"variable_product_options_inner\"]/div[4]/div[2]/h3/select")));				
						WebElement Sel = driver.findElement(By.xpath("//*[@id='variable_product_options_inner']/div[5]/div[1]/h3/select"));
						Thread.sleep(1567);
						Select Select = new Select(Sel);	
						Select.selectByVisibleText(String_word);
						Thread.sleep(2000);
						//		String di = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div[3]/div[1]/div[6]/form/div/div/div[1]/div[3]/div/div[2]/div/div[7]/div/div[4]/div/h3/strong")).getText();
						driver.findElement(By.xpath("//*[@id='variable_product_options_inner']/div[5]/div[1]/h3/strong")).click();
						driver.findElement(By.xpath("//*[@id='variable_regular_price_"+y+"']")).sendKeys(ReggularPric[y]); 
						driver.findElement(By.xpath("//*[@id='variable_weight"+y+"']")).sendKeys(Weight[y]); 
						driver.findElement(By.xpath("//*[@id='variable_wcj_msrp_"+y+"']")).sendKeys(SalePric[y]);
						driver.findElement(By.xpath("//*[@id='inspector-text-control-"+y+"']")).sendKeys(DelivaryNote);
						WebElement comission = driver.findElement(By.xpath("//*[@id=\"variable_product_options_inner\"]/div[5]/div/div/div/div[8]/p/input"));
						comission.sendKeys(Commissio[y]);
						driver.findElement(By.xpath("//button[text()='Save changes']")).click();		
						JavascriptExecutor executor1 = (JavascriptExecutor) driver;			
						executor1.executeScript("arguments[0].scrollIntoView();", elem);     
						Thread.sleep(2500);
					}	y++;			        	
				}wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='variable_product_options_inner']/div[1]/div[1]/select")));
				WebElement cel = driver.findElement(By.xpath("//*[@id='variable_product_options_inner']/div[1]/div[1]/select"));
				Select cele = new Select(cel);
				cele.selectByVisibleText(one);
				} catch (Exception e) {
					System.out.println("Variations");
					e.printStackTrace();
				}
			WebElement savedraft = driver.findElement(By.xpath("//*[@id='save-post']"));
			savedraft.click();			
			System.out.println(id+" IS Completed");
			
		}
	}
}