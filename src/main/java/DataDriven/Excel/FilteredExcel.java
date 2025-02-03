package DataDriven.Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;


public class FilteredExcel {

	
	public static void main(String[] args) {
		String filePath = "F:\\AllProjectResources\\Excel\\testdata\\ProxyDataSheet.xlsx";
		
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");  
        LocalDateTime now = LocalDateTime.now(); 
        
		String outputFilePath = "F:\\AllProjectResources\\Excel\\testdata\\Output_ProxyDataSheet"+dtf(format(now))+".xlsx";
		
		 
	        
	        System.out.println("Current Date and Time: " + dtf.format(now));
		try (FileInputStream fis = new FileInputStream(filePath);
				Workbook workbook = new XSSFWorkbook(fis);
				FileOutputStream fos = new FileOutputStream(outputFilePath)) {

			Sheet originalSheet = workbook.getSheetAt(0); // Assuming data is in the first sheet
			Sheet filteredSheet = workbook.createSheet("Filtered_Data");

			int rowCount = 0;
			for (Row row : originalSheet) {
				Cell skipCell = row.getCell(6); // Assuming "Skip" is the 7th column (index 6)

				if (rowCount == 0 || (skipCell != null && "N".equalsIgnoreCase(skipCell.getStringCellValue()))) {
					Row newRow = filteredSheet.createRow(rowCount);

					for (int i = 0; i < row.getLastCellNum(); i++) {
						Cell originalCell = row.getCell(i);
						Cell newCell = newRow.createCell(i);

						if (originalCell != null) {
							switch (originalCell.getCellType()) {
							case STRING:
								newCell.setCellValue(originalCell.getStringCellValue());
								break;
							case NUMERIC:
								newCell.setCellValue(originalCell.getNumericCellValue());
								break;
							case BOOLEAN:
								newCell.setCellValue(originalCell.getBooleanCellValue());
								break;
							default:
								newCell.setCellValue("");
							}
						}
					}
					rowCount++;
				}
			}

			workbook.write(fos);
			System.out.println("Filtered Excel file created successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static String dtf(Object format) {
		// TODO Auto-generated method stub
		return null;
	}

	private static Object format(LocalDateTime now) {
		// TODO Auto-generated method stub
		return null;
	}

}

