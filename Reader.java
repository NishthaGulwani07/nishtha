package DataDriven;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class Reader {

	private static void getdata() throws IOException {
              
		File s= new File("C:\\Users\\HP\\eclipse-workspace\\Cucumber\\DataFiles\\data.xlsx");
		FileInputStream fi=new FileInputStream(s);
				Workbook wb =  new  XSSFWorkbook(fi);
		Sheet sheet = wb.getSheet("Data");
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(0);
		CellType ce = cell.getCellType();
		if (ce.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
			
		}else if (ce.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int a =(int) numericCellValue;
			System.out.println(a);
			
		}}
	public static void AllDatas() throws IOException {
		File s = new File("C:\\Users\\HP\\eclipse-workspace\\Cucumber\\DataFiles\\data.xlsx");
		FileInputStream ss = new FileInputStream(s);
		Workbook sss = new XSSFWorkbook(ss);
		Sheet t = sss.getSheet("Data");
		int Rows = t.getPhysicalNumberOfRows();
		for (int i = 0; i < Rows; i++) {
			Row row = t.getRow(i);
			int Cells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < Cells; j++) {
				Cell cell = row.getCell(j);
				CellType CT = cell.getCellType();
				if (CT.equals(CellType.STRING)) {
					String s1 = cell.getStringCellValue();
					System.out.println(s1);
				} else if (CT.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int r = (int) numericCellValue;
					System.out.print(r);
				}
				System.out.print("||");
			}
			System.out.println();

		}

	}
	public static void main(String[] args) throws IOException {
//  getdata();
  AllDatas();
	}

}
