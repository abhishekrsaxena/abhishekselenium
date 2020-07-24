package Selflearning.Mavenjava;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class datadriven {

	
	public ArrayList<String> getdata(String TestCaseName) throws IOException
	{
		ArrayList<String> a= new ArrayList<String>();
		
		
		FileInputStream fis=new FileInputStream("C://work//DemoData.xlsx");
		
			XSSFWorkbook workbook=new XSSFWorkbook(fis);
			
			int sheets= workbook.getNumberOfSheets();
			
			for(int i=0;i<sheets;i++)
			{
				if (workbook.getSheetName(i).equalsIgnoreCase("TestData"))
						{
					XSSFSheet sheet= workbook.getSheetAt(i);
					Iterator<Row>  rows= sheet.iterator();// sheet is collection of row
					Row firstrow = rows.next();
					Iterator<Cell> ce= firstrow.cellIterator();//row is collection of cell
					
					int k=0;
					int coloumn =0;
					while(ce.hasNext())
					{
					Cell value= ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases"))
					{
							coloumn =k;
					}
					      k++;
					      
					}
					// System.out.println(coloumn);
					while(rows.hasNext())
					{
						Row r= rows.next();
						if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(TestCaseName))
						{
							Iterator<Cell >cv=r.cellIterator();
							while(cv.hasNext())
							{
								
								Cell c=cv.next();
								if(c.getCellType()== CellType.STRING)
								{
									a.add(c.getStringCellValue());
								}
								else {
									a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								
							}
							
								
								
							}
						}
						
						
					}
					 
					 
					}
				
								
			}
			return a;
			
	
			
			
	}
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
	}

}
