package Task13;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook book = new XSSFWorkbook ("C:\\Users\\USER\\eclipse-workspace\\TestExcel\\src\\main\\java\\Task13\\Task13Excel.xlsx");
		XSSFSheet sheet = book .getSheetAt(0);
		
		int rowcount = sheet .getLastRowNum();
		int columncount = sheet .getRow(0).getLastCellNum();
		
		String [][] data = new String [rowcount][columncount];
		//Get into row 
	for (int i=1 ; i<= rowcount ; i++) {
		XSSFRow row= sheet.getRow(i);
		//Get into cell
		for (int j=0 ; j<columncount;j++) {
			XSSFCell cell = row.getCell(j);
			//To read data from excel
			data [i-1][j] = cell.getStringCellValue();
			System.out.println(cell.getStringCellValue());
		}
	}
	book.close();

}
}
