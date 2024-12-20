package exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcelForLoop {

	public static void main(String[] args) throws IOException {

		String excelFilePath = ".\\datafiles\\countries.xlsx";
		FileInputStream inputstream = new FileInputStream(excelFilePath);

		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		// read data using for loop

		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();

		for(int r=0;r<=rows;r++)
		{
			XSSFRow row = sheet.getRow(r);

			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);

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
