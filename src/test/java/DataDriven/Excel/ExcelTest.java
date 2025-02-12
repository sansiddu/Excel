package DataDriven.Excel;


import java.util.HashSet;
import java.util.Iterator;


import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.util.List;
import java.util.Set;


public class ExcelTest {

	@BeforeTest
	public void setup() {
		ExcelUtility.filterExcelData(); // Ensure filtered data is generated before tests
	}

	@DataProvider(name = "filteredDataProvider")
	public static Iterator<Object[]> filteredDataProvider() {
		return ExcelUtility.getFilteredData();
	}

	@Test(dataProvider = "filteredDataProvider")
	public void printFilteredData(Object... rowData) {
		for (Object data : rowData) {
			System.out.print(data + " || ");
		}
		System.out.println();
	}

	@Test(dataProvider = "filteredDataProvider")
	public void Extractdata(String SlNo, String TestCase, String Scenario, String Login_UserName, String Proxy_Type, String IN_AccountName, 
			String Skip, String O_final_Result, String Date, String Time, String IN_UpcomingVisits, String O_UPcomingVisits, String Up_Comments,
			String IN_PastVisits, String O_PastVisits, String Past_Comments, String IN_ScheduledAppointment, String O_ScheduledAppointment,
			String ScheduledAppointment_Comments, String IN_V) {

		List<String> accountNames = ExcelUtility.getINAccountNames();
		System.out.println("Extracted IN_AccountNames: " + accountNames);
		 Set<String> processedNames = new HashSet<>();
		 
		for (String accountName : accountNames) {
			if(!processedNames.contains(accountName)) {
			switch (accountName) {
			case "Parent-Teena":

				System.out.println("Validation Passed: Parent-Teena exists");
				System.out.println("Proxy Parent-k: " + Proxy_Type +"  ");
				System.out.println("Proxy Parent-k: " + Skip+"  ");
				System.out.println("Proxy Parent-k: " + IN_UpcomingVisits+"  ");
				break;
			case "Parent-k":
				System.out.println("Validation Passed: Parent-k exists");
				break;
			case "Parent-N":
				System.out.println("Validation Passed: Parent-N exists");
				break;
			case "Parent-O":
				System.out.println("Validation Passed: Parent-O exists");
				break;  
			default:
				System.out.println("No specific validation for ");
			}
			processedNames.add(accountName);
		}
	}
}
}

//		    @Test(dataProvider = "filteredDataProvider")
//		    public void fillform(String SlNo, String TestCase, String Scenario, String Login_UserName,   String Proxy_Type,
//		    		String IN_AccountName, String IN_UpcomingVisits, String Skip, String IN_PastVisits, String IN_ScheduledAppointment) {
//		     
//		    	for (Object data : rowData) {
//		            System.out.print(data + " || ");
//		        }
//		        System.out.println("Proxy: " + Proxy_Type +"  ");
//		        System.out.println("Proxy: " + IN_AccountName+"  ");
//		        System.out.println("Proxy: " + IN_UpcomingVisits+"  ");
//		     
//		        if("Parent-k".contentEquals(IN_AccountName)) {
//		        	System.out.println("Proxy Parent-k: " + Proxy_Type +"  ");
//			        System.out.println("Proxy Parent-k: " + Skip+"  ");
//			        System.out.println("Proxy Parent-k: " + IN_UpcomingVisits+"  ");
//		        }


//	@Test(dataProvider = "filteredDataProvider")
//	public List<String> testINAccountNames(String SlNo, String TestCase, String Scenario, String Login_UserName, String Proxy_Type, String IN_AccountName, 
//			String Skip, String O_final_Result, String Date, String Time, String IN_UpcomingVisits, String O_UPcomingVisits, String Up_Comments,
//			String IN_PastVisits, String O_PastVisits, String Past_Comments, String IN_ScheduledAppointment, String O_ScheduledAppointment,
//			String ScheduledAppointment_Comments, String IN_V) {
//		
//		List<String> accountNames = ExcelUtility.getINAccountNames();
//		System.out.println("Extracted IN_AccountNames: " + accountNames);
//
//		for (String accountName : accountNames) {
//			switch (accountName) {
//			case "Parent-Teena":
//				System.out.println("Performing action for Parent-Teena");
//				System.out.println("ProxyTeena: " + Proxy_Type +"  ");
//				System.out.println("ProxyTeena: " + IN_AccountName+"  ");
//				System.out.println("ProxyTeena: " + IN_UpcomingVisits+"  ");
//				// Add required action here
//				break;
//			case "Parent":
//				System.out.println("Performing action for Parent");
//				System.out.println("ProxyParent: " + Proxy_Type +"  ");
//				System.out.println("ProxyParent: " + IN_AccountName+"  ");
//				System.out.println("ProxyParent: " + IN_UpcomingVisits+"  ");
//				// Add required action here
//				break;
//			case "Parent-k":
//				System.out.println("Performing action for Parent-K");
//				
//				// Add required action here
//				break;
//			case "Parent-N":
//				System.out.println("Performing action for Parent-N");
//				// Add required action here
//				break;
//			case "Parent-O":
//				System.out.println("Performing action for Parent-O");
//				// Add required action here
//				break;
//			default:
//				System.out.println("No specific action for " + accountName);
//			}
//		}
//		return accountNames;
//	}

//	@Test
//	public void useTestINAccountNames() {
//		List<String> accountNames = testINAccountNames();
//		System.out.println("Processing extracted account names in another method: " + accountNames);
//
//		for (String accountName : accountNames) {
//			System.out.println("Handling " + accountName + " in a different logic");
//		}







