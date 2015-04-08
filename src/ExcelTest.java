import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelTest {

	static XSSFRow row;
	static XSSFCell cell;

	public static void main(String[] args) throws IOException, InterruptedException {

			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("Mine");
			row = sheet.createRow(0);
			row.createCell(0).setCellValue("번호");
			row.createCell(1).setCellValue("이름");
			row.createCell(2).setCellValue("점수");

			FileOutputStream outFile;
			try {
				outFile = new FileOutputStream("c:/excel/test2.xlsx");
				try {
					wb.write(outFile);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				try {
					outFile.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		
		



		/*		String excelFile = "C:/excel/workbook.xlsx";

		@SuppressWarnings("resource")
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		Sheet sheet = (Sheet) wb.createSheet("First Sheet");
		if(sheet == null) {
			System.out.println("Sheet1 is Null");
			return;
		}*/

	}
	/*try {

		InputStream is = new FileInputStream("C:/excel/test.xlsx");
		Workbook wb = WorkbookFactory.create(is);
		Sheet sheet = (Sheet) wb.getSheetAt(0);
		Row row = ((Cell) sheet).getRow();
		Cell cell = row.getCell(0);
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}


	}*/
}
