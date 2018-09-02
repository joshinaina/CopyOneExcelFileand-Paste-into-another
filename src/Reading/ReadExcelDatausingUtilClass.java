package Reading;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDatausingUtilClass 
{
		public static void main(String args[]) throws Exception {
			ExcelApiTest eat = new ExcelApiTest("C:\\Users\\vinod\\Desktop\\vinods.xlsx");
			
			for(int i=0;i<=15;i++)
				
			{
				String data0=eat.getCellData("Sheet2", 0, i);
				System.out.println(("value of Sheet 2" +data0));
			
				
				//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
                
				FileInputStream fis1 = new FileInputStream("C:\\Users\\vinod\\Desktop\\sunainas.xlsx");
            
				XSSFWorkbook workbook = new XSSFWorkbook(fis1);
            //call the getSheet() method of Workbook and pass the Sheet Name here. 
            //In this case I have given the sheet name as “TestData” 
               //or if you use the method getSheetAt(), you can pass sheet number starting from 0. Index starts with 0.
        //    XSSFSheet sheet = workbook.getSheet("Input");
            XSSFSheet sheet = workbook.getSheet("Input");
            //Now create a row number and a cell where we want to enter a value. 
            //Here im about to write my test data in the cell B2. It reads Column B as 1 and Row 2 as 1. Column and Row values start from 0.
            //The below line of code will search for row number 2 and column number 2 (i.e., B) and will create a space. 
               //The createCell() method is present inside Row class.
              
            
            
            
           Row row2 = sheet.createRow(0);
                    Cell cell = row2.createCell(i);
            //Now we need to find out the type of the value we want to enter. 
               //If it is a string, we need to set the cell type as string 
               //if it is numeric, we need to set the cell type as number
            cell.setCellType(cell.CELL_TYPE_STRING);
            cell.setCellValue(data0);
        
            
    FileOutputStream fos = new FileOutputStream("C:\\Users\\vinod\\Desktop\\sunainas.xlsx");
    
    workbook.write(fos);
	//	fos.close();
		System.out.println("END OF WRITING DATA IN EXCEL");
		}
	}
	
	
		
		

				
				
				
				
				
			
			
			
			
			
			
			//System.out.println(eat.getCellData("Sheet2", 0, 0));
			//System.out.println(eat.getCellData("Sheet2", 0, 1));
			//System.out.println(eat.getCellData("Sheet2", 0, 2));
			/*System.out.println(eat.getCellData("Sheet2", 0, 3));
			System.out.println(eat.getCellData("Sheet2", 0, 4));
			System.out.println(eat.getCellData("Sheet2", 0, 5));
			System.out.println(eat.getCellData("Sheet2", 0, 6));
			System.out.println(eat.getCellData("Sheet2", 0, 7));
			System.out.println(eat.getCellData("Sheet2", 0, 8));
			System.out.println(eat.getCellData("Sheet2", 0, 9));
			System.out.println(eat.getCellData("Sheet2", 0, 10));
			System.out.println(eat.getCellData("Sheet2", 0, 11));
			System.out.println(eat.getCellData("Sheet2", 0, 12));
			System.out.println(eat.getCellData("Sheet2", 0, 13));
			System.out.println(eat.getCellData("Sheet2", 0, 14));
			System.out.println(eat.getCellData("Sheet2", 0, 15));
			*/
			
			/*//SHeet Joy
			
			System.out.println("************joyfileData***************");
			System.out.println(eat.getCellData("joy", 0, 0));
			System.out.println(eat.getCellData("joy", 0, 1));
			System.out.println(eat.getCellData("joy", 0, 2));
			System.out.println(eat.getCellData("joy", 0, 3));
			System.out.println(eat.getCellData("joy", 0, 4));
			System.out.println(eat.getCellData("joy", 0, 5));
			//first row
			System.out.println("***************************");
			System.out.println(eat.getCellData("joy", 1, 0));
			System.out.println(eat.getCellData("joy", 1, 1));
			System.out.println(eat.getCellData("joy", 1, 2));
			System.out.println(eat.getCellData("joy", 1, 3));
			System.out.println(eat.getCellData("joy", 1, 4));
			System.out.println(eat.getCellData("joy", 1, 5));
			
			//second row
			System.out.println("***************************");
			System.out.println(eat.getCellData("joy", 2, 0));
			System.out.println(eat.getCellData("joy", 2, 1));
			System.out.println(eat.getCellData("joy", 2, 2));
			System.out.println(eat.getCellData("joy", 2, 3));
			System.out.println(eat.getCellData("joy", 2, 4));
			System.out.println(eat.getCellData("joy", 2, 5));
			
			//sheet Darlene

			
			System.out.println("************DarlenefileData***************");
			System.out.println(eat.getCellData("Darlene", 0, 0));
			System.out.println(eat.getCellData("Darlene", 0, 1));
			System.out.println(eat.getCellData("Darlene", 0, 2));
			System.out.println(eat.getCellData("Darlene", 0, 3));
			System.out.println(eat.getCellData("Darlene", 0, 4));
			System.out.println(eat.getCellData("Darlene", 0, 5));
			System.out.println(eat.getCellData("Darlene", 0, 6));
			System.out.println(eat.getCellData("Darlene", 0, 7));
			System.out.println(eat.getCellData("Darlene", 0, 8));
			System.out.println(eat.getCellData("Darlene", 0, 9));
			System.out.println(eat.getCellData("Darlene", 0, 10));
			System.out.println(eat.getCellData("Darlene", 0, 11));
			System.out.println(eat.getCellData("Darlene", 0, 12));
			System.out.println(eat.getCellData("Darlene", 0, 13));
			System.out.println(eat.getCellData("Darlene", 0, 14));
			System.out.println(eat.getCellData("Darlene", 0, 15));
			System.out.println(eat.getCellData("Darlene", 0, 16));
			System.out.println(eat.getCellData("Darlene", 0, 17));
			
			//SHeet Sunaina
			System.out.println("************SunainafileData***************");
			System.out.println(eat.getCellData("Sunaina", 0, 0));
			System.out.println(eat.getCellData("Sunaina", 0, 1));
			System.out.println(eat.getCellData("Sunaina", 0, 2));
			System.out.println(eat.getCellData("Sunaina", 0, 3));
			System.out.println(eat.getCellData("Sunaina", 0, 4));
			System.out.println(eat.getCellData("Sunaina", 0, 5));
			System.out.println(eat.getCellData("Sunaina", 0, 6));
			System.out.println(eat.getCellData("Sunaina", 0, 7));
			System.out.println(eat.getCellData("Sunaina", 0, 8));
			System.out.println(eat.getCellData("Sunaina", 0, 9));
			
			//sheet2
			
			System.out.println(eat.getCellData("Sunaina", 1, 0));
			System.out.println(eat.getCellData("Sunaina", 1, 1));
			System.out.println(eat.getCellData("Sunaina", 1, 2));
			System.out.println(eat.getCellData("Sunaina", 1, 3));
			System.out.println(eat.getCellData("Sunaina", 1, 4));
			System.out.println(eat.getCellData("Sunaina", 1, 5));
			System.out.println(eat.getCellData("Sunaina", 1, 6));
			System.out.println(eat.getCellData("Sunaina", 1, 7));
			System.out.println(eat.getCellData("Sunaina", 1, 8));
			System.out.println(eat.getCellData("Sunaina", 1, 9));
			
			
			*/
			
			
		
		}
		
		
		
		
		
		


	


			
		
	
		






