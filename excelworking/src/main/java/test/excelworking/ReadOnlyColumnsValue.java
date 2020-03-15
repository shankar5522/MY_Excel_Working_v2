package test.excelworking;
/**
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author Shankar
 * 
 * This class is only use to read a particular column values based on Row number
 * Also, we see some work-around related to sheets as well.
 *
 */
public class ReadOnlyColumnsValue {

	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		File file = null;
		FileInputStream fileInputStream = null;
		XSSFWorkbook workbook = null;
		XSSFSheet sheet = null;
		XSSFRow rows = null;
		XSSFCell cell = null;

		String pathName = System.getProperty("user.dir") + "/data/myData.xlsx";

		file = new File(pathName);
		fileInputStream = new FileInputStream(file);
		workbook = new XSSFWorkbook(fileInputStream);

		//Not of Sheet in Workbook
		//Print all the sheet name and Index
		int no_Sheets = workbook.getNumberOfSheets();
		System.out.println("No of Sheets in WorkBook " + no_Sheets);
		String sheetName;
		for (int i = 0; i < no_Sheets; i++) {
			sheetName = workbook.getSheetName(i);
			System.out.println("Index " + i + " has sheet Name " + sheetName);
		}

		// No Of Rows in Sheet
		sheet = workbook.getSheet("input");
		int no_of_rows = sheet.getLastRowNum() + 1;
		System.out.println("No of Rows has in Sheet : " + no_of_rows);

		//No of Col in Sheet
		rows = sheet.getRow(0);
		int no_of_columns = rows.getLastCellNum();
		System.out.println("No of Columns in the sheet : " + no_of_columns);

		//get the data from a specified row
		System.out.println("----------------GET THE DATA FROM SPECIFED ROW------------------");
		int cellType = 0;
		for (int i = 0; i < no_of_columns; i++) {
			cell = sheet.getRow(3).getCell(i); // here Roe is fixed

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
				System.out.print(i + " columns Has differnt cell type");
			}
		}
		System.out.println();

		System.out.println("----------------GET THE DATA FROM SPECIFED COLUMNS------------------");
		//get the data from specified columns
		DateFormat dateFormat = null;
		Date date = null;
		for (int i = 0; i < no_of_rows; i++) {
			cell = sheet.getRow(i).getCell(5); // here columns is fixed

			cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_NUMERIC) {
				System.out.println(cell.getNumericCellValue() + "			");
			} else if (cellType == Cell.CELL_TYPE_STRING) {
				System.out.println(cell.getStringCellValue() + "			");
			} else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
				System.out.println(cell.getBooleanCellValue() + "			");
			} else if (cellType == Cell.CELL_TYPE_BLANK) {
				System.out.println(cell.getStringCellValue());
			} else if (HSSFDateUtil.isCellDateFormatted(cell)) {
				dateFormat = new SimpleDateFormat("dd/mm/yyyy");
				date = cell.getDateCellValue();
				System.out.println(date);
				System.out.println(dateFormat.format(date));
			} else {
				System.out.println(i + " columns Has differnt cell type");
			}
		}
		System.out.println();

	}

}
