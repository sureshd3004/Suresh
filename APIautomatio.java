package books;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.testng.annotations.Test;
import okhttp3.OkHttpClient;
import okhttp3.Request;

public class APIautomatio  {	

	@Test
	public void status() throws InterruptedException, IOException {
		String path="C:\\Users\\sures\\OneDrive\\Desktop\\sana.xlsx";
		for (int j = 1; j < 20000; j++) {
			System.out.println(j);
			OkHttpClient client = new OkHttpClient();
			Request request = new Request.Builder()
					.url("https://my.cloudtalk.io/api/contacts/index.json?page="+j+"&limit=1000")
					.get()
					.addHeader("Authorization", "Basic Auth Key neddeddddddddddddddddddd=")
					.build();
			okhttp3.Response response = client.newCall(request).execute();
			String jsonResponse = response.body().string();
			//	System.out.println(jsonResponse);
			JSONObject jsonObject = new JSONObject(jsonResponse);
			JSONArray dataArray = jsonObject.getJSONObject("responseData").getJSONArray("data");

			FileInputStream fis = new FileInputStream(path);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Create header row
			Row headerRow = sheet.createRow(0);
			String[] headers = {"ID", "Name", "Company", "Phone Number", "Email"};
			for (int i = 0; i < headers.length; i++) {
				headerRow.createCell(i).setCellValue(headers[i]);
			}

			// Iterate over JSON data and write to Excel
			for (int i = 0; i < dataArray.length(); i++) {
				JSONObject contact = dataArray.getJSONObject(i);
				JSONObject contactData = contact.getJSONObject("Contact");
			//	JSONObject contactNumberr = contact.getJSONObject("ContactNumber");
				System.out.println(contactData);
				int num = sheet.getLastRowNum();
				Row row = sheet.createRow(1+num);
				row.createCell(0).setCellValue(contactData != null ? contactData.optString("company") : "");
				row.createCell(1).setCellValue(contactData.getString("name"));
				try{
					row.createCell(2).setCellValue(contactData.getString("company"));
				}catch (org.json.JSONException e) {
					row.createCell(2).setCellValue("null");
				}
				//	row.createCell(2).setCellValue(contactData.getString("company"));
				JSONObject contactNumber = contact.optJSONObject("ContactNumber");
				row.createCell(3).setCellValue(contactNumber != null ? contactNumber.optString("public_number") : "");
				row.createCell(6).setCellValue(contactNumber != null ? contactNumber.optString("country_code_id") : "");
				row.createCell(7).setCellValue(contactNumber != null ? contactNumber.optString("id") : "");
				JSONObject contactEmail = contact.optJSONObject("ContactEmail");
				row.createCell(4).setCellValue(contactEmail != null ? contactEmail.optString("email") : "");
				row.createCell(5).setCellValue(contactEmail != null ? contactEmail.optString("contact_id") : "");
				
			}
			FileOutputStream fileOut = new FileOutputStream(path);
			workbook.write(fileOut);

		} 
	}
}
