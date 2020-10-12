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

public class Excel {

	/*public static void main(String args[]) throws IOException
	{*/
	
	 public ArrayList<String> getData(String testcaseName) throws IOException {
	 ArrayList<String> a=new ArrayList<String>();
	 
		FileInputStream fis=new FileInputStream("C:\\Users\\TECHOLUTION\\Desktop\\excel file.xlsx");
		XSSFWorkbook w=new XSSFWorkbook(fis);
	      int sheets=w.getNumberOfSheets();
	      System.out.println(sheets);
	     String sheetname= w.getSheetName(0);
	     System.out.println(sheetname);
		
		for (int i=0;i<sheets;i++)
		{
			
	     if(w.getSheetName(i).equalsIgnoreCase("Sheet1"))
	     {
	    	XSSFSheet sheet= w.getSheetAt(i);
	    	System.out.println(sheet);
	    	Iterator<Row> rows=sheet.iterator();
	    	Row firstrow=rows.next();
	    	Iterator<Cell> cells=firstrow.cellIterator();
	    	int k=0;
	    	int column=0;
	    	while(cells.hasNext())
	    	{
	    		Cell column_values=cells.next();
	    		if(column_values.getStringCellValue().equalsIgnoreCase("Testcases"))
	    		{
	    			column=k;
	    		}
	    		k++;
	    		}
	    	System.out.println(column);
	    	while(rows.hasNext())
	    	{
	    		Row r=rows.next();
	    		if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
	    		{
	    			Iterator<Cell> cv=r.cellIterator();
	    			while(cv.hasNext())
	    			{
	    				//System.out.println(cv.next().getStringCellValue());
	    				//a.add(cv.next().getStringCellValue());
	    				Cell c=cv.next();
	    				if(c.getCellTypeEnum()==CellType.STRING)
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
		return a;
	}

	/*
	 * public static void main(String args[]) {
	 * 
	 * }
	 */
}
//}
