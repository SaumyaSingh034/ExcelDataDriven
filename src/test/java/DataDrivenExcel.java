import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDrivenExcel {

public ArrayList<String> getDataFromExcel(String TestCaseName) throws IOException
{
	ArrayList<String> a = new ArrayList<String>();
	FileInputStream file = new FileInputStream("D:\\Selenium\\ExcelDOc\\DataTest.xlsx");
	
	XSSFWorkbook workbook = new XSSFWorkbook(file);
	
	int sheetCount = workbook.getNumberOfSheets();
	for(int i=0;i<sheetCount;i++)
	{
		if(workbook.getSheetName(i).equalsIgnoreCase("TestData"))
		{
			XSSFSheet sheet = workbook.getSheetAt(i);
			//Identify Test Case Row
			Iterator<Row> rows = sheet.iterator();
			Row firstRow = rows.next();
			Iterator<Cell> ce = firstRow.cellIterator();
			int column = 0;
			int k = 0;
			while(ce.hasNext())
			{
				Cell value = ce.next();
				if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
				{
					column = k;
				}
				k++;
			}
			System.out.println(column);
			while(rows.hasNext())
			{
				Row r = rows.next();
				if(r.getCell(column).getStringCellValue().equalsIgnoreCase(TestCaseName))
				{
					Iterator<Cell> cv = r.cellIterator();
					while(cv.hasNext())
					{
						Cell c = cv.next();
						a.add(c.getStringCellValue());
					}
				}
			}
			
			
		}
		
		
	}
	return a;
}
}

