package a3;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;


public class ExcelTest {

	static Row row;

	public static void main(String[] args) throws IOException, InterruptedException {

	    
//        sxssf.writeWorkbook("sxssf-sample.xlsx");

	    
	    Test sxssf = new Test(Test.SXSSF);
	    sxssf.process("c:/excel/test.xlsx");
	}
}
