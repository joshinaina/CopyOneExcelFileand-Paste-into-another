package Reading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApiTest {
	
	public FileInputStream fis=null;
	public XSSFWorkbook workbook=null;
	public POIXMLDocumentPart sheet =null;
	public XSSFRow row = null;
	public XSSFCell cell=null;
	
	public ExcelApiTest(String xlFilePath) throws IOException
	{
		fis = new FileInputStream(xlFilePath);
		workbook = new XSSFWorkbook(fis);
		fis.close();
		
		
		
		
	}
	
	public String getCellData(String sheetName, int colNum,int rowNum) 
	{
		try 
		{
			sheet = workbook.getSheet(sheetName);
			row = ((XSSFSheet) sheet).getRow(rowNum);
			cell = row.getCell(colNum);
			
			if(cell.getCellTypeEnum()==CellType.STRING)
				
				return cell.getStringCellValue();
			
			else if(cell.getCellTypeEnum()==CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				
				String cellValue = String.valueOf(cell.getNumericCellValue());
				if(HSSFDateUtil.isCellDateFormatted(cell))
				{
					DateFormat dt = new SimpleDateFormat("dd/MM/yy");
					Date date = (Date) cell.getDateCellValue();
					cellValue = dt.format(date);
				}
				return cellValue;
			}else if(cell.getCellTypeEnum()==CellType.BLANK)
				return"";
			else
				return String.valueOf(cell.getBooleanCellValue());
		}
		catch (Exception e)
		{
			e.printStackTrace();
			return "No Matched valve";
		}
				
					
					
					
					
					
				}
			}
		
	
		


