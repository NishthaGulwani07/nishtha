package DataDriven;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XLUtility {

	
	private static final String Data = null;
	private static final String IndexedColors = null;
	private static final String FillPatternType = null;
	public FileInputStream fi;
	public FileOutputStream fo;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFRow cell;
	public CellStyle style;
	String path=null;
	
	XLUtility(String path){
		
		this.path=path;
		}
	
	public int getRowCount(String sheetName) throws IOException{
		
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(Data);
		int rowCount=sheet.getLastRowNum();
		((Closeable) workbook).close();
		fi.close();
		return rowCount;
		
	}
	
	public int getCellCount(String sheetName,int rownum) throws IOException{
		
		fi=new FileInputStream(path);
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		int cellcount=row.getLastCellNum();
		((Closeable) workbook).close();
		fi.close();
		return cellcount;
	}
	
	public String getCellData(String sheetName,int rownum,int colnum) throws IOException{
		
		fi=new FileInputStream(path);
		
		
		workbook =new XSSFWorkbook(fi);
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(rownum);
		
		cell=row.getCell(colnum);
		return sheetName;
	}

public void setCellData(String sheetName,int rownum,int colnum, String data) throws IOException{
	
	fi=new FileInputStream(path);
	workbook =new XSSFWorkbook(fi);
	sheet=workbook.getSheet(sheetName);
	row=sheet.getRow(rownum);
	
	cell=row.createCell(colnum);
	cell.setCellValue(data);
	
	fo=new FileOutputStream(path);
	workbook.write(fo);
	((Closeable) workbook).close();
	fi.close();
	fo.close();
	
	
	
}

}