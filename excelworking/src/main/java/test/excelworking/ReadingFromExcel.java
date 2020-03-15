package test.excelworking;
/**
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Shankar
 *
 */
public class ReadingFromExcel {

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

		sheet = workbook.getSheet("input");

		// get the number of rows in sheet
		int numberOfRows = sheet.getLastRowNum() + 1;
		System.out.println("Number of Rows in Sheet : " + numberOfRows);

		// Here zero because want to calculate the number of rows in that sheet
		rows = sheet.getRow(0);
		int numberOfCols = rows.getLastCellNum();
		System.out.println("Number Of Cols in Sheet : " + numberOfCols);

		//to print only headings
		for (int i = 0; i < numberOfCols; i++) {
			cell = rows.getCell(i);
			System.out.print(cell.getStringCellValue() + " ");
		}
		System.out.println();

		int cellType = 0;
		// To get the all the rows and columns values
		for (int i = 1; i < numberOfRows; i++) {
			rows = sheet.getRow(i);
			for (int j = 0; j < numberOfCols; j++) {
				//System.out.println(rows.getCell(j).getStringCellValue() + " ");
				cell = rows.getCell(j);
				cellType = cell.getCellType();
				if (cellType == Cell.CELL_TYPE_NUMERIC) {
					System.out.print(cell.getNumericCellValue() + "			");
				} else if (cellType == Cell.CELL_TYPE_STRING) {
					System.out.print(cell.getStringCellValue() + "			");
				} else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
					System.out.print(cell.getBooleanCellValue() + "			");
				} else if (cellType == Cell.CELL_TYPE_BLANK) {
					System.out.println(cell.getStringCellValue());
				} else {
					System.out.print(i + " " + j + "Has differnt cell type");
				}
			}
			System.out.println();

			//Once done the excel process, don't forget to close the connection to avoide the Memory leackage.
			workbook.close();
			fis.close();
			System.out.println("Done with Writting");

		}

	}//Main

}//class
