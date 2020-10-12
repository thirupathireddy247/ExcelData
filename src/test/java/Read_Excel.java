import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Excel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
		  FileInputStream fis=new
		  FileInputStream("C:\\\\Users\\\\TECHOLUTION\\\\Desktop\\\\excel file.xlsx");
		  XSSFWorkbook w=new XSSFWorkbook(fis); XSSFSheet sheet=w.getSheetAt(1);
		  System.out.println(sheet);
		  
		  
		  XSSFRow rows=sheet.getRow(1); System.out.println(rows); XSSFCell
		  cells=rows.getCell(2); System.out.println(cells);
		  
		  
		/* * System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
		 * System.out.println(sheet.getRow(0).getCell(1).getStringCellValue());
		 * System.out.println(sheet.getRow(0).getCell(2).getStringCellValue());
		 System.out.println(sheet.getRow(1).getCell(2,10).getNumericCellValue());
		 
		FileInputStream fis=new FileInputStream("C:\\\\Users\\\\TECHOLUTION\\\\Desktop\\\\excel file.xlsx");
				  XSSFWorkbook w=new XSSFWorkbook(fis); 
				  String sheet1=w.getSheetName(0);*/
			
				  
		

	}

}
