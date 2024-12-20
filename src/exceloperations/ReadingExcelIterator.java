package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelIterator {

	public static void main(String[] args) throws IOException {
		
		String excelfilepath = ".\\datafiles\\countries.xlsx";
		FileInputStream stream = new FileInputStream(excelfilepath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(stream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		// read data using iterator method
		
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext())
			{
				Cell cell = cellIterator.next();
				switch(cell.getCellType())
				{
				case STRING: 
					System.out.print(cell.getStringCellValue()); 
					break;

				case NUMERIC: 
					System.out.print(cell.getNumericCellValue()); 
					break;

				case BOOLEAN: 
					System.out.print(cell.getBooleanCellValue()); 
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}

	}

}
