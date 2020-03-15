package test.excelworking;
/**
 * 
 */


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Shankar
 * 
 * Set the major or minor as per the user age
 *
 */
public class WriteIntoFile {

	static File file = null;
	static FileInputStream fis = null;
	static XSSFWorkbook workbook = null;
	static XSSFSheet sheet = null;
	static XSSFRow rows = null;
	static XSSFCell cell = null;

	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//Windows path format - > C:\Users\Shankar\eclipse-workspace-oxygen\Excel_Working\data
		String pathname = System.getProperty("user.dir") + "/data/myData.xlsx";

		file = new File(pathname);
		fis = new FileInputStream(file);
		workbook = new XSSFWorkbook(fis);

		sheet = workbook.getSheet("operation");

		// get the number of rows in sheet
		int numberOfRows = sheet.getLastRowNum() + 1;
		System.out.println("Number of Rows in Sheet : " + numberOfRows);

		// Here zero because want to calculate the number of rows in that sheet
		rows = sheet.getRow(0);
		int numberOfCols = rows.getLastCellNum();
		System.out.println("Number Of Cols in Sheet : " + numberOfCols);

		int cellType = 0;
		double age = 0.0;
		// To get the all the rows and columns values
		for (int i = 1; i < numberOfRows; i++) {
			cell = sheet.getRow(i).getCell(0);
			cellType = cell.getCellType();

			if (cellType == Cell.CELL_TYPE_NUMERIC) {
				age = cell.getNumericCellValue();
				System.out.print("Your Age is " + age);
			} else if (cellType == Cell.CELL_TYPE_STRING) {
				System.out.print(cell.getStringCellValue());
			} else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
				System.out.print(cell.getBooleanCellValue());
			} else if (cellType == Cell.CELL_TYPE_BLANK) {
				System.out.print(cell.getStringCellValue());
			} else {
				System.out.print(i + " Rows with Columns 0 Has differnt cell type");
			}
			System.out.println();

			if (age >= 18) {
				sheet.getRow(i).getCell(1).setCellValue("MAJOR");
			} else {
				sheet.getRow(i).getCell(1).setCellValue("Minor");
			}
		} //for loop

		//Once done the excel process, don't forget to close the connection to avoide the Memory leackage.
		fis.close();

		FileOutputStream fileOutputStream = new FileOutputStream(pathname);
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();

	}//main
}//class
