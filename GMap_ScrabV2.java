package books;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.Duration;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GMap_ScrabV2{

	@Test
	public  void main() throws IOException, InterruptedException {

		String search = "Flavours manufacturers in Delhi";

		String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\Output.xlsx";
		WebDriver driver = new ChromeDriver();
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
					String mobileNumberPattern = "(\\+\\d{1,3})?\\s*\\d{5}\\s*\\d{5}";

					Pattern pattern = Pattern.compile(mobileNumberPattern);
					Matcher matcher = pattern.matcher(one);

					String key = "1000.422c263cbc3cf035ce4a6e5feff99062.096c4ef5b819a1b1729927b558442f9f";
					String jsonResponse=null;
					String id=null;
					String phone=null;
					String email=null;
					if (two.length()<50) { 
						try {
							 URL urls = new URL(two);
						      HttpURLConnection connection = (HttpURLConnection) urls.openConnection();
						      connection.setRequestMethod("GET");

						      BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
						      StringBuilder stringBuilder = new StringBuilder();
						      String line;

						      while ((line = reader.readLine()) != null) {
						        stringBuilder.append(line);
						      }

						      reader.close();
						      String pageContent = stringBuilder.toString();

						      // Consider a more robust email regex (search online)
						      Pattern emailPattern = Pattern.compile("[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}");

						      // Find emails in the page content using regex
						      Matcher matcher2 = emailPattern.matcher(pageContent);
						      while (matcher2.find()) {
						        email = matcher2.group();
						        System.out.println("Email found: " + email);
						      }}
						catch (Exception e) {        }
					}
						if (matcher.find()) {
							phone = matcher.group();
							System.out.println("Mobile Number: " + phone);
							//	String title = cell1.getStringCellValue();
						}
						//lead
						try {
							OkHttpClient client = new OkHttpClient();
							Request request = new Request.Builder()
									.url("https://www.zohoapis.com/crm/v6/Leads/search?criteria=Phone:equals:" + phone)
									.get()
									.addHeader("Authorization", "Zoho-oauthtoken "+key)
									.build();
							Response response = client.newCall(request).execute();
							jsonResponse = response.body().string();
							System.out.println("json lead is ="+jsonResponse);

							if (response.isSuccessful()) {
								JSONObject jsonObject = new JSONObject(jsonResponse);
								JSONArray itemsArray = jsonObject.getJSONArray("data");
								for (int k = 0; i < itemsArray.length(); k++) {
									//  JSONObject jsonObject = new JSONObject(jsonResponse);
									JSONArray dataArray = jsonObject.getJSONArray("data");
									JSONObject firstDataObject = dataArray.getJSONObject(0);
									JSONObject ownerObject = firstDataObject.getJSONObject("Owner");

									id = ownerObject.getString("id");
								}		System.out.println(id);
							}} catch (Exception e) {
							}

						//contact
						try {
							OkHttpClient client = new OkHttpClient();
							Request request = new Request.Builder()
									.url("https://www.zohoapis.com/crm/v6/Contacts/search?criteria=Phone:equals:" + phone )
									.get()
									.addHeader("Authorization", "Zoho-oauthtoken "+key)
									.build();
							okhttp3.Response response = client.newCall(request).execute();
							jsonResponse = response.body().string();
							System.out.println("json is contact is="+jsonResponse);
							if (response.isSuccessful()) {
								JSONObject jsonObject = new JSONObject(jsonResponse);
								JSONArray itemsArray = jsonObject.getJSONArray("data");
								for (int k = 0; k < itemsArray.length(); k++) {
									JSONObject dataObject = itemsArray.getJSONObject(k);
									String ids = dataObject.getString("id");
									id =ids+" Account";
								}		
							}} catch (Exception e) {
							}
						//account
						try {
							OkHttpClient client = new OkHttpClient();
							Request request = new Request.Builder()
									.url("https://www.zohoapis.com/crm/v6/Accounts/search?criteria=Phone:equals:" + phone )
									.get()
									.addHeader("Authorization", "Zoho-oauthtoken "+key)
									.build();
							okhttp3.Response response = client.newCall(request).execute();
							jsonResponse = response.body().string();
							System.out.println("json is acc is ="+jsonResponse);
							if (response.isSuccessful()) {
								JSONObject jsonObject = new JSONObject(jsonResponse);
								JSONArray itemsArray = jsonObject.getJSONArray("data");
								for (int k = 0; k < itemsArray.length(); k++) {
									JSONObject dataObject = itemsArray.getJSONObject(k);
									String ids = dataObject.getString("id");
									id =ids+" Account";
								}		
							}} catch (Exception e) {
					} 
					FileInputStream fis = new FileInputStream(wpath);
					XSSFWorkbook wbook = new XSSFWorkbook(fis);

					XSSFSheet sheet = wbook.getSheetAt(0);
					Row row = sheet.createRow(b++);

					org.apache.poi.ss.usermodel.Cell cell1 = row.createCell(1, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell2 = row.createCell(2, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell3 = row.createCell(3, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell4 = row.createCell(4, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell5 = row.createCell(5, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell6 = row.createCell(6, CellType.STRING);
					org.apache.poi.ss.usermodel.Cell cell7 = row.createCell(7, CellType.STRING);
					cell1.setCellValue(zero);
					cell2.setCellValue(one);
					cell3.setCellValue(phone);
					cell4.setCellValue(two);
					cell5.setCellValue(id);
					cell6.setCellValue(email);
					cell7.setCellValue(jsonResponse);
					FileOutputStream fos = new FileOutputStream(wpath);
					wbook.write(fos);
					fis.close();
					fos.close();				
				} b++;
			}
			WebElement page = driver.findElement(By.xpath("//*[@id='rl_ist0']/div/div[2]/div/table/tbody/tr/td["+c+++"]/a"));
			Thread.sleep(2345);
			wait.until(ExpectedConditions.elementToBeClickable(page));
			page.click();
		}	
	}}