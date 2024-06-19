package com12taste;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import com.neverbounce.api.client.NeverbounceClient;
import com.neverbounce.api.client.NeverbounceClientFactory;
import com.neverbounce.api.model.Result;
import com.neverbounce.api.model.SingleCheckResponse;

public class NeverBounce {

	public static void main(String[] args) throws Exception {

		String rpath="C:\\Users\\sures\\OneDrive\\Desktop\\Read.xlsx";

		for (int i=1; i <8; i++){
			System.out.println(i);
			XSSFWorkbook rbook = new XSSFWorkbook(rpath);
			XSSFSheet sheet = rbook.getSheetAt(0);
			XSSFRow li = sheet.getRow(i);
			XSSFCell cell = li.getCell(0);
			String value = cell.getStringCellValue();
			String token = "private_Api Key";

			NeverbounceClient neverbounceClient = NeverbounceClientFactory.create(token);
			SingleCheckResponse singleCheckResponse = neverbounceClient
					.prepareSingleCheckRequest()
					.withEmail(value) // address to verify
					.withAddressInfo(true)  // return address info with response
					.withCreditsInfo(true)  // return account credits info with response
					.withTimeout(20)  // only wait on slow email servers for 20 seconds max
					.build()
					.execute();
			rbook.close();	
			JSONObject jsonObject = new JSONObject(singleCheckResponse);
			//   JSONArray employees = jsonObject.getJSONArray("result");
			Result result = singleCheckResponse.getResult();
			// JsonUtils.printJson(result);
			String hi = result.toString().replace("\"", "");
			String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\write.xlsx";     
			FileInputStream fis = new FileInputStream(wpath);
			XSSFWorkbook wbook = new XSSFWorkbook(fis);
			XSSFSheet shee = wbook.getSheet("Sheet3");
			//	String final = result.toString();
			Row ro = shee.createRow(i);
			org.apache.poi.ss.usermodel.Cell cell1 = ro.createCell(1);
			org.apache.poi.ss.usermodel.Cell cell2 = ro.createCell(2);
			cell1.setCellValue(value);
			cell2.setCellValue(hi);
			System.out.println(hi);
			FileOutputStream fos = new FileOutputStream(wpath);
			wbook.write(fos);
			fis.close();
			fos.close();
			wbook.close();

		}
	}
}
