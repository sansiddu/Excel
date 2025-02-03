package DataDriven.Excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class OutputExceldataprovider {

	public static void main(String[] args) {
		String filePath = "ProxyDataSheet.xlsx";
		String outputFilePath = "Filtered_ProxyDataSheet.xlsx";

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

	@DataProvider(name = "filteredDataProvider")
	public static Iterator<Object[]> filteredDataProvider() {
		List<Object[]> testData = new ArrayList<>();
		String filePath = "Filtered_ProxyDataSheet.xlsx";

		try (FileInputStream fis = new FileInputStream(filePath);
				Workbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet = workbook.getSheet("Filtered_Data");
			for (Row row : sheet) {
				List<Object> rowData = new ArrayList<>();
				for (Cell cell : row) {
					switch (cell.getCellType()) {
					case STRING:
						rowData.add(cell.getStringCellValue());
						break;
					case NUMERIC:
						rowData.add(cell.getNumericCellValue());
						break;
					case BOOLEAN:
						rowData.add(cell.getBooleanCellValue());
						break;
					default:
						rowData.add("");
					}
				}
				testData.add(rowData.toArray());
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return testData.iterator();
	}

	@Test(dataProvider = "filteredDataProvider")
	public void printFilteredData(Object... rowData) {
		for (Object data : rowData) {
			System.out.print(data + " | ");
		}
		System.out.println();
	}
}
 
