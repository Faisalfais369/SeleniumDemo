package dringDataFromExcelUsingDataProvider;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

//for using excel data  here we have to use Apache poi and  Apache Poi-ooxml dependencies

public class ExcelDataUsingDataProvider {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider="ExceldatausingDataProvider")
	public void getDataFromExcel(String value1, String value2, String value3)
	{
		System.out.println(value1+" "+value2+" "+value3);
	}
	
	
	@DataProvider(name="ExceldatausingDataProvider") 
	public Object[][] giveDatafromExcel() throws IOException
	{
		FileInputStream fis = new FileInputStream("C:/Users/Ahmed Faisal/ExcelForTesting/Book1.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		
		int sheetCount = workBook.getNumberOfSheets();
		
		XSSFSheet sheet = null;
		for(int i = 0 ; i < sheetCount ; i++)
		{

			if(workBook.getSheetAt(i).getSheetName().equalsIgnoreCase("Numbertwo"))
			{
				sheet = workBook.getSheetAt(i);
			}
		}
				
				int rowCount = sheet.getPhysicalNumberOfRows();
				XSSFRow firstRow = sheet.getRow(0);
				int columnCount =  firstRow.getLastCellNum();
				
				Object data[][] = new Object[rowCount][columnCount];
				
				
				for(int j = 0 ; j < rowCount ; j++)
				{
					Row row = sheet.getRow(j);
				
					for(int k = 0 ;  k < columnCount ; k++)
					{
						Cell cell = row.getCell(k);
						data[j][k] = formatter.formatCellValue(cell);
					}
					
				}
				workBook.close();
				return data;					
	}	
}
