package drivingExcelData;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FirstExcelProgram {
	public ArrayList<String> getData(String path) throws IOException
	{
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(path);
		
		XSSFWorkbook workBook = new XSSFWorkbook(fis);//accessing the ExcelSheet
		
		int sheetCount = workBook.getNumberOfSheets();//under Excel sheet finding out number of sheets present inside the Excel
		
		int k = 0, column = 0 ;
		
		for(int i  = 0 ; i < sheetCount ; i++)
		{
			if(workBook.getSheetName(i).contains("Numberone"))//checking whether the excel sheet "Numberone" is present or not, if yes then go  inside
				//this sheet
			{
				XSSFSheet sheet = workBook.getSheetAt(i);//getting the access of "Numberone" sheet
				
				Iterator<Row> row = sheet.rowIterator();//implementing  the Row iterator in the sheet too check the "Test cases" row
				Row firstRow = row.next();// taking the access for first row
				
				Iterator<Cell> cell = firstRow.cellIterator();//implementing the cell iterator in a row to check the "Test cases", to check
				//which column has Purchase Order text, to take the data present in that row
				
				while(cell.hasNext())
				{
					Cell value = cell.next();//Going through Each cell/column
					if(value.getStringCellValue().contains("test cases"))
					{
						column = k;//taking the column number
					}
					k++;
				}
				
				while(row.hasNext())
				{
					Row absoluteRow = row.next();//after the taking the column then going into the seconRow 
					Cell absoluteCell = absoluteRow.getCell(column);
					if(absoluteCell.getStringCellValue().contains("Purchase Order"))
					{
						Iterator<Cell> cel = absoluteRow.cellIterator();
						while(cel.hasNext())
						{
							Cell c = cel.next();
							if(c.getCellType() == CellType.STRING)
							{
								a.add(c.getStringCellValue());
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		workBook.close();
		return a;
		
	}
	public  static void main(String args[])
	{
		
	}

}
