package scrab;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UploadSheetIN {

	public static void main(String[] args) {

		//		try {
		String rpath="C:\\Users\\sures\\OneDrive\\Desktop\\EU 18_12 A (1).xlsx";
		String wpath="C:\\Users\\sures\\OneDrive\\Desktop\\write.xlsx";
		try 
		(FileInputStream sourceFileInputStream = new FileInputStream(rpath);
				FileInputStream destinationFileInputStream = new FileInputStream(wpath);
				XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);
				XSSFWorkbook destinationWorkbook = new XSSFWorkbook(destinationFileInputStream)) {
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
			XSSFSheet destinationSheet = destinationWorkbook.getSheetAt(0);
			int lastRowNum = sourceSheet.getLastRowNum();
			for (int i = 1; i <= lastRowNum; i++) {
				String a7 = "";
				String a77 = "";
				String a8= "";
				String a9 = "";
				System.out.println(i);
				XSSFRow sourceRow = sourceSheet.getRow(i);
				XSSFRow destinationRow = destinationSheet.createRow(i);
				if (sourceRow != null) {
					XSSFCell sourceCell = sourceRow.getCell(2);	
					XSSFCell lastname = sourceRow.getCell(3);               
					XSSFCell email = sourceRow.getCell(4);  
					XSSFCell source0 = sourceRow.getCell(6);  
					XSSFCell source = sourceRow.getCell(7);  
					XSSFCell source1 = sourceRow.getCell(9);  
					XSSFCell source2= sourceRow.getCell(10);  
					XSSFCell source3= sourceRow.getCell(12);  
					XSSFCell source4= sourceRow.getCell(13);  
					XSSFCell source5= sourceRow.getCell(14);  
					XSSFCell source6= sourceRow.getCell(15);  
					XSSFCell source7= sourceRow.getCell(16);  
					XSSFCell source77= sourceRow.getCell(17);  
					//   System.out.println(source7.getStringCellValue());
					XSSFCell source8= sourceRow.getCell(18);  
					XSSFCell source9= sourceRow.getCell(19);  
					XSSFCell source10= sourceRow.getCell(20);  
					XSSFCell source11= sourceRow.getCell(21);  
					XSSFCell source12= sourceRow.getCell(41);  
					XSSFCell source13= sourceRow.getCell(39); 
					//	XSSFCell source14= sourceRow.getCell(19); 
					XSSFCellStyle style=destinationWorkbook.createCellStyle(); 
					style.setFillBackgroundColor(IndexedColors.RED.getIndex()); 
			
					XSSFCell Name = destinationRow.createCell(0,CellType.STRING);
					XSSFCell lName = destinationRow.createCell(1,CellType.STRING);
					XSSFCell mail = destinationRow.createCell(2,CellType.STRING);     
					XSSFCell mailq = destinationRow.createCell(3,CellType.STRING); 
					XSSFCell lsorce = destinationRow.createCell(4,CellType.STRING); 
					XSSFCell mail0 = destinationRow.createCell(5,CellType.STRING); 
					XSSFCell mail1 = destinationRow.createCell(6,CellType.STRING);
					XSSFCell mail2 = destinationRow.createCell(7,CellType.STRING);
					XSSFCell mail3 = destinationRow.createCell(8,CellType.STRING);
					XSSFCell mail4 = destinationRow.createCell(9,CellType.STRING);
					XSSFCell mail5 = destinationRow.createCell(10,CellType.STRING);
					XSSFCell mail6 = destinationRow.createCell(11,CellType.STRING);
					//mail6.setCellStyle();
					XSSFCell mail7 = destinationRow.createCell(12,CellType.STRING);
					XSSFCell mail8 = destinationRow.createCell(13,CellType.STRING);
					XSSFCell mail9 = destinationRow.createCell(14,CellType.STRING);
					XSSFCell mail10 = destinationRow.createCell(15,CellType.STRING);
					XSSFCell mail11 = destinationRow.createCell(16,CellType.STRING);
					XSSFCell mail12 = destinationRow.createCell(17,CellType.STRING);
					XSSFCell mail13 = destinationRow.createCell(18,CellType.STRING);
					XSSFCell mail14 = destinationRow.createCell(19,CellType.STRING);
					XSSFCell mail15 = destinationRow.createCell(20,CellType.STRING);
					XSSFCell mail16 = destinationRow.createCell(21,CellType.STRING);
					XSSFCell mail17 = destinationRow.createCell(22,CellType.STRING);
					mail17.setCellValue("EN");
					mail16.setCellValue("IN");;

					// Copy the cell value to the destination cell
					String val1 = sourceCell.getStringCellValue();
					String val2 = lastname.getStringCellValue();
					String val3 = email.getStringCellValue();
					String val4 = source0.getStringCellValue();
					String val5 = source.getStringCellValue();
					try {										
						String val6 = source1.getStringCellValue();
						mail0.setCellValue(val6);
					} catch (Exception e) {							
					}
					String val7 = source2.getStringCellValue();
					String val8 = source3.getStringCellValue();
					String val9 = source4.getStringCellValue();

					try {
						String val10= source5.getStringCellValue();
						String val11= source6.getStringCellValue();	
						mail4.setCellValue(val10);
						mail5.setCellValue(val11);

					} catch (Exception e) {
						System.out.println("Missing State or City In = "+i);
					}
					try {
						a7 = source7.getStringCellValue();
					} catch (Exception e) {
					}
					try {
						a77 = source77.getStringCellValue();
						mail7.setCellValue(a77);
					} catch (Exception e) {
					}
					try {
						a8 = source8.getStringCellValue();
						mail8.setCellValue(a8);
					} catch (Exception e) {
					}try {
						a9 = source9.getStringCellValue();
						mail9.setCellValue(a9);
					} catch (Exception e) {					
					}
					if (a7.isEmpty()) {
						//		 mail6.setCellValue(source7.getStringCellValue());// This line is not necessary but included for clarity

						if (a7.isEmpty() && !a77.isEmpty()) {
							mail6.setCellValue(a77);
						}

						// Copy values from Column 3 to Column 1 if Column 1 is empty and Column 2 is empty
						if (a7.isEmpty() && a77.isEmpty() && !a8.isEmpty()) {
							mail6.setCellValue(a8);
						}

						// Copy values from Column 4 to Column 1 if Column 1 is empty, Column 2 is empty, and Column 3 is empty
						if (a7.isEmpty() && a77.isEmpty() && a8.isEmpty() && !a9.isEmpty()) {
							mail6.setCellValue(a9);
						}
					}
					String val15= source10.getStringCellValue();
					String val16= source11.getRawValue();
					String val17= source12.getRawValue();
					String val18= source13.getStringCellValue();
					//		String val19 =source14.getStringCellValue();
					Name.setCellValue(val1);
					lName.setCellValue(val2);
					mail.setCellValue(val3);
					mailq.setCellValue(val4);
					lsorce.setCellValue(val5);						
					mail1.setCellValue(val7);
					mail2.setCellValue(val8);
					mail3.setCellValue(val9);
					mail15.setCellValue(val15);
					mail10.setCellValue(val16);
					mail11.setCellValue(val17);
					mail12.setCellValue(val18);
					mail13.setCellValue("iliyas.r@12taste.com");					

					double num = source11.getNumericCellValue();
					if (num >=1 && num <=50 ) {	
						mail14.setCellValue("Small Account");
					}
					if (num >=51 && num <=1000) {		 mail14.setCellValue("Mid Account");				
					}
					if (num >=1001 && num <=5000) {		 mail14.setCellValue("Large Account");				
					}
					if (a7.isEmpty() && a77.isEmpty() && a8.isEmpty() && a9.isEmpty()) {
						mail6.setCellValue("No Contact Is Found");
						mail6.setCellStyle(style);
					}		if (!a7.isEmpty()) {
						mail6.setCellValue(a7);
						if (a7.contains("TPS")) {
							mail6.setCellValue("No Contact Is Found");
							mail6.setCellStyle(style);
						}
						if (a7.contains("DNC")) {
							mail6.setCellValue("No Contact Is Found");
							mail6.setCellStyle(style);
						}
					}
				}      
				try (FileOutputStream fileOutputStream = new FileOutputStream(wpath)) {
					destinationWorkbook.write(fileOutputStream);
				}}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}//finally {}
}