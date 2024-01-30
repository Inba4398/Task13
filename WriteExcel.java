package Task13;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {
		// 1. To create a excel book
		XSSFWorkbook book = new XSSFWorkbook ();
		// 2. To create a  sheet
		XSSFSheet sheet = book .createSheet("Sheet1");
		//3.Write data on excel file
		Object[][] data = {
				{"Name" , "Age", "Email"} ,
				{"John Doe" , "30" ,"john@test.com"},
				{"Jane Doe", "28", "john@test.com"},
				{"Bob Smith" , "35" ,"jacky@example.co"},
				{"Swapnil" , "37" ,"swapnil@example.com"},
		};
		//Create row
		int rowCount = 0;
		
		for (Object[] row1 : data) {
			XSSFRow row = sheet.createRow(rowCount++);
			
			//Create column
			int columnCount =0;
			
			for (Object col : row1) {
				XSSFCell cell = row . createCell (columnCount++); 
				
				if (col instanceof String) {
					cell.setCellValue((String)col);
				}else if (col instanceof String) {
					cell.setCellValue((Integer) col);
				}
				}
				}
		try (
				FileOutputStream output = new FileOutputStream ("C:\\Users\\USER\\eclipse-workspace\\TestExcel\\src\\main\\java\\Task13\\Task13Excel.xlsx");) {
			     book.write(output);
		
		
	}


			
	

		
					
				
		
		
		
		

	}

}
