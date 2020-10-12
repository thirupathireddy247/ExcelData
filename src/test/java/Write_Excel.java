import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Excel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream fis=new FileInputStream("C:\\Users\\TECHOLUTION\\Desktop\\excel file.xlsx");
		XSSFWorkbook book=new XSSFWorkbook(fis);
		XSSFSheet sheet=book.getSheetAt(1);
		sheet.getRow(0).createCell(0).setCellValue("thirupathi");
		sheet.getRow(1).createCell(1).setCellValue("goapl");
		sheet.getRow(1).createCell(2).setCellValue("barathi");
		sheet.getRow(2).createCell(1).setCellValue(9441);
		sheet.getRow(2).createCell(2).setCellValue(4265);
		
		FileOutputStream fos=new FileOutputStream("C:\\\\Users\\\\TECHOLUTION\\\\Desktop\\\\excel file - Copy.xlsx");
		book.write(fos);
		}

}
