package DataDriven.Excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class ExcelUtility {

	// Excel Utility Class
	    public static final String INPUT_FILE = "F:\\AllProjectResources\\Excel\\testdata\\ProxyDataSheet.xlsx";
	    public static final String OUTPUT_FILE = "F:\\AllProjectResources\\Excel\\testdata\\OrignalProxyDataSheet.xlsx";
	    
	    public static void filterExcelData() {
	        try (FileInputStream fis = new FileInputStream(INPUT_FILE);
	             Workbook workbook = new XSSFWorkbook(fis);
	             FileOutputStream fos = new FileOutputStream(OUTPUT_FILE)) {
	            
	            Sheet originalSheet = workbook.getSheetAt(0);
	            Sheet filteredSheet = workbook.createSheet("Filtered_Data");
	            
	            int rowCount = 0;
	            for (Row row : originalSheet) {
	                Cell skipCell = row.getCell(6);
	                
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
//	                                case NUMERIC:
//	                                    newCell.setCellValue(originalCell.getNumericCellValue());
//	                                    break;
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
	    
	    public static Iterator<Object[]> getFilteredData() {
	        List<Object[]> testData = new ArrayList<>();
	        
	        try (FileInputStream fis = new FileInputStream(OUTPUT_FILE);
	             Workbook workbook = new XSSFWorkbook(fis)) {
	            
	            Sheet sheet = workbook.getSheet("Filtered_Data");
	            for (Row row : sheet) {
	                List<Object> rowData = new ArrayList<>();
	                for (Cell cell : row) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            rowData.add(cell.getStringCellValue());
	                            break;
//	                        case NUMERIC:
//	                            rowData.add(cell.getNumericCellValue());
//	                            break;
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
	    
	    public static List<String> getINAccountNames() {
	        List<String> accountNames = new ArrayList<>();
	        
	        try (FileInputStream fis = new FileInputStream(OUTPUT_FILE);
	             Workbook workbook = new XSSFWorkbook(fis)) {
	            
	            Sheet sheet = workbook.getSheet("Filtered_Data");
	            int accountNameColumn = -1;
	            
	            for (Row row : sheet) {
	                if (accountNameColumn == -1) {
	                    // Identify the column index of "IN_AccountName"
	                    for (Cell cell : row) {
	                        if (cell.getCellType() == CellType.STRING && "IN_AccountName".equalsIgnoreCase(cell.getStringCellValue())) {
	                            accountNameColumn = cell.getColumnIndex();
	                            break;
	                        }
	                    }
	                    continue;
	                }
	                
	                if (accountNameColumn != -1) {
	                    Cell accountNameCell = row.getCell(accountNameColumn);
	                    if (accountNameCell != null && accountNameCell.getCellType() == CellType.STRING) {
	                        accountNames.add(accountNameCell.getStringCellValue());
	                    }
	                }
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        return accountNames;
	    }
	}
