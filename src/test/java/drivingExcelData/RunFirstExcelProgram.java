package drivingExcelData;

import java.io.IOException;
import java.util.ArrayList;

public class RunFirstExcelProgram {
	
	public static void main(String[] args) throws IOException
	{
		FirstExcelProgram f  = new FirstExcelProgram();
		ArrayList<String> a = f.getData("C:/Users/Ahmed Faisal/ExcelForTesting/Book1.xlsx");
		
		System.out.println(a.get(0));
		System.out.println(a.get(1));
		System.out.println(a.get(2));
		System.out.println(a.get(3));
		System.out.println(a.get(4));
	}

}
