import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader{
	private static Workbook wb;
	private static Sheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static Row row;
	private static Cell cell;
	public static void main(String[] args) throws Exception {
		
		fis = new FileInputStream("C:\\Users\\golla\\eclipse-workspace\\ApachePOI\\pivot_table.xls");
		wb = WorkbookFactory.create(fis);
		sh=wb.getSheet("Voters");
		int noOfRows=sh.getLastRowNum();
		System.out.println(noOfRows);
		
		for (int i=3980;i<=noOfRows;i++) {
			System.out.println(sh.getRow(i).getCell(0));
		}
/*		
 * https://qavalidation.com/2015/03/selenium-excel-read-and-write-apachepoi.html/
 * row = sh.createRow(4002);
		cell= row.createCell(0);
		cell.setCellValue("mani");
		System.out.println(cell.getStringCellValue());
		fos = new FileOutputStream("C:\\Users\\golla\\eclipse-workspace\\ApachePOI\\pivot_table.xls");
		wb.write(fos);
		fos.flush();
		fos.close();
		System.out.println("Done");*/
	
	}
}
