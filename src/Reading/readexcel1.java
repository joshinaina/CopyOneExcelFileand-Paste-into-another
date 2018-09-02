package Reading;
import org.apache.commons.io.FileUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.google.common.collect.Table.Cell;

public class readexcel1 {
	

	
	int i;
	int j;
	String data0;
	public FileInputStream fis=null;
	public XSSFWorkbook workbook=null;
	public POIXMLDocumentPart sheet =null;
	public XSSFRow row = null;
	public XSSFCell cell=null;
		
	
	public void read() throws IOException
	{
File src=new File("C:\\Users\\vinod\\Desktop\\vinods.xlsx");
		
		FileInputStream fis =new FileInputStream(src);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		XSSFSheet sheet1=wb.getSheetAt(0);
		
		int rowcount=sheet1.getLastRowNum();
		
		for(i=0;i<=19;i++)
			
		{
			for(j=0;j<=8;j++)
			{
				
				 data0=sheet1.getRow(i).getCell(j).getStringCellValue();
				System.out.println("Data from Row " +i+" is "+data0);
			
			
			
				
			//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
                
					FileInputStream fis1 = new FileInputStream("C:\\Users\\vinod\\Desktop\\sunainas.xlsx");
                
					XSSFWorkbook workbook = new XSSFWorkbook(fis1);
                //call the getSheet() method of Workbook and pass the Sheet Name here. 
                //In this case I have given the sheet name as “TestData” 
                   //or if you use the method getSheetAt(), you can pass sheet number starting from 0. Index starts with 0.
            //    XSSFSheet sheet = workbook.getSheet("Input");
                XSSFSheet sheet = workbook.getSheet("Input");
                Row row=sheet.createRow(0);
                
                Cell cell=row.createCell(i);
                
                cell.setCellValue(data0);
                		
                //Now create a row number and a cell where we want to enter a value. 
                //Here im about to write my test data in the cell B2. It reads Column B as 1 and Row 2 as 1. Column and Row values start from 0.
                //The below line of code will search for row number 2 and column number 2 (i.e., B) and will create a space. 
                   //The createCell() method is present inside Row class.
                  
                
                
                
               /*Row row2 = sheet.createRow(i);
                        Cell cell = row2.createCell(j);
                //Now we need to find out the type of the value we want to enter. 
                   //If it is a string, we need to set the cell type as string 
                   //if it is numeric, we need to set the cell type as number
                cell.setCellType(cell.CELL_TYPE_STRING);
                cell.setCellValue(data0);
            */
                
        FileOutputStream fos = new FileOutputStream("C:\\Users\\vinod\\Desktop\\sunainas.xlsx");
        
        workbook.write(fos);
		//	fos.close();
			System.out.println("END OF WRITING DATA IN EXCEL");
			}
		}
		
		
			}
			
		
	
		
				
				
			
			public static void main(String[] args) throws IOException {
				// TODO Auto-generated method stub
				
				readexcel1 obj=new readexcel1();
				obj.read();

						
				
				}
			
			}
		
	
		
		
		
			


		
	
