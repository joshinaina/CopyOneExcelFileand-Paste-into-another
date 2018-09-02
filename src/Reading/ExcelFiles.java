package Reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gargoylesoftware.htmlunit.javascript.host.Iterator;

public class ExcelFiles {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		File src=new File("C:\\Users\\vinod\\Desktop\\testing.xlsx");
		
		FileInputStream fis =new FileInputStream(src);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		XSSFSheet sheet1=wb.getSheetAt(0);
		//for print First row First cell
		
		String data0=sheet1.getRow(0).getCell(1).getStringCellValue();
		System.out.println("Data From Excel is " +data0);
		// next line printed
		String data3=sheet1.getRow(1).getCell(19).getStringCellValue();
		System.out.println("Data From Excel is " +data3);
		
		//For Print First row Second cell
		
		String data1=sheet1.getRow(8).getCell(19).getStringCellValue();
		System.out.println("Data From Excel is "+data1);
		
		// for print first row 3rd cell
		String data2=sheet1.getRow(9).getCell(19).getStringCellValue();
		System.out.println("Data From Excel is "+data2);
		
		
		wb.close();
		
		
		
		
		
		
	
		
		
		
	}
}
		
		
		