package Reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NickM {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File src=new File("C:\\Users\\vinod\\Desktop\\vinods.xlsx");
		
		FileInputStream fis =new FileInputStream(src);
		
		XSSFWorkbook wb =new XSSFWorkbook(fis);
		 XSSFSheet sheet2=wb.getSheetAt(1);
		 
		  String data1=sheet2.getRow(0).getCell(0).getStringCellValue();
		  
		 System.out.println("Data From Excel is "+data1);
		
		//second coloumn
		 
		String data2=sheet2.getRow(0).getCell(1).getStringCellValue();
		  
		 System.out.println("Data From Excel is "+data2);
		 
double data3=sheet2.getRow(1).getCell(1).getNumericCellValue();
		  
		 //System.out.println("Data From Second Row "+data3);
		
		

	}

}
