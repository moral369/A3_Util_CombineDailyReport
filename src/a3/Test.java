package a3;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public class Test {
	private Sheet s;
	public static final int HSSF = 0;
	public static final int XSSF = 1;
	public static final int SXSSF = 2;
	private Workbook wb;


	public Test(int xssf2) {
		// TODO Auto-generated constructor stub
		/*switch(xssf2){
        case HSSF : wb = new HSSFWorkbook();
            break;
        case XSSF : wb = new XSSFWorkbook();
            break;
        case SXSSF : wb = new SXSSFWorkbook(-1);
            break;
		}*/
	}

	public Row getRow(int i){
		Row r = s.getRow(i);
		if(r==null)
			r = s.createRow(i);
		return r;
	}

	public Cell getCell(int row,int cell){
		Row r = getRow(row);
		Cell c = r.getCell(cell);
		if(c==null)
			c = r.createCell(cell);
		return c;
	}

	public void setCellValue(int row, int cell, String cellvalue){
		Cell c = getCell(row,cell);
		c.setCellValue(cellvalue);
	}

	public void writeWorkbook(String fileName){
		long start = System.currentTimeMillis();
		try{
			s = wb.createSheet("sample Sheet");
			for(int i=0;i<10000;i++){
				setCellValue(i,0,"Test_Title_"+i);
				setCellValue(i,1,"Test_Title_"+i);
				setCellValue(i,2,"Test_Title_"+i);
				setCellValue(i,3,"Test_Title_"+i);
				setCellValue(i,4,"Test_Title_"+i);
				setCellValue(i,5,"Test_Title_"+i);
				setCellValue(i,6,"Test_Title_"+i);
				setCellValue(i,7,"Test_Title_"+i);
				setCellValue(i,8,"Test_Title_"+i);
				setCellValue(i,9,"Test_Title_"+i);
				setCellValue(i,10,"Test_Title_"+i);
			}
			wb.write(new FileOutputStream(fileName));
		}catch(Exception e){
			e.printStackTrace();
			System.err.println(e.getMessage());
		}
		long end = System.currentTimeMillis();
		System.out.println("writeHSSFWorkbook : "+(end-start));
	}

	// TODO Auto-generated method stub
	public void process(String string) throws IOException, InvalidFormatException {

		Workbook inWb = new SXSSFWorkbook(200);
		FileInputStream fs = new FileInputStream("c:/excel/test.xlsx");
		wb =WorkbookFactory.create(fs);
		
		Sheet sheet = wb.getSheetAt(0);
		
		Workbook outWb = new SXSSFWorkbook(200);
		
		FileOutputStream outFile;
		outFile = new FileOutputStream("output.xlsx");
		outWb.write(outFile);
		http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-excel
		
	}
}
