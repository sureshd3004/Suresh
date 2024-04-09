package books;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.testng.annotations.Test;
import okhttp3.OkHttpClient;
import okhttp3.Request;

public class APIautomati  {	

	@Test
	public void status() throws InterruptedException, IOException {
		String path="C:\\Users\\sures\\OneDrive\\Desktop\\Read.xlsx";
		for (int j = 1; j < 71; j++) {
			System.out.println(j);
			OkHttpClient client = new OkHttpClient();
			Request request = new Request.Builder()
					.url("https://www.zohoapis.com/books/v3/items?page="+j+"&per_page=200")
					.get()
					.addHeader("Authorization", "Zoho-oauthtoken 1000.598b1b5c094ae6c44acbefa4cf12db69.7395576f9d7c3927f8d2488196039f80")
					.build();
			okhttp3.Response response = client.newCall(request).execute();
			String jsonResponse = response.body().string();
			//	System.out.println(jsonResponse);
			JSONObject jsonObject = new JSONObject(jsonResponse);
			JSONArray itemsArray = jsonObject.getJSONArray("items");
			FileInputStream fis = new FileInputStream(path);
			XSSFWorkbook wbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = wbook.getSheetAt(0);
			// Write data from JSON array to Excel
			int lastRow = sheet.getLastRowNum();
			for (int i = 0; i < itemsArray.length(); i++) {
				JSONObject itemObject = itemsArray.getJSONObject(i);
				XSSFRow row = sheet.createRow(lastRow+i + 1);
				if (itemObject.has("item_id")) {
					row.createCell(0).setCellValue(itemObject.getString("item_id"));
				} if (itemObject.has("name")) {
					row.createCell(1).setCellValue(itemObject.getString("name"));}
				if (itemObject.has("item_name")) {
					row.createCell(2).setCellValue(itemObject.getString("item_name"));}
				if (itemObject.has("status")) {
					row.createCell(3).setCellValue(itemObject.getString("status"));}
				if (itemObject.has("cf_product_url")) {
					row.createCell(4).setCellValue(itemObject.getString("cf_product_url"));}
				if (itemObject.has("cf_product_url_unfor0matted")) {
					row.createCell(5).setCellValue(itemObject.getString("cf_product_url_unfor0matted"));
				}
			} FileOutputStream fileOut = new FileOutputStream(path);
			wbook.write(fileOut);
			wbook.close();
		}
	}
}