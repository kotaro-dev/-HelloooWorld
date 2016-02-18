package mylocal.disk01.study.main;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample4 {

	private static final String BOOK_PATH = "C:\\Dev\\";
	private static final String BOOK_NAME = "Sample4.xlsx";
	private static final String SHEET_NAME = "mySheet";

	public static void main(String[] args) {
		// TODO 自動生成されたメソッド・スタブ
		new Sample4().execute();
//		new Sample4().executeCreateBook();
		new Sample4().executeUpdateBook();
	}

	private void execute() {
		System.out.println("Sample4 Control Excel");
	}

	private void executeCreateBook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet(SHEET_NAME);

		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		//set cell value
		cell.setCellValue("Hello World! 2!");

		//style set
		CellStyle style = book.createCellStyle();
		Font font = book.createFont();
		font.setColor(IndexedColors.AQUA.getIndex());

		style.setFont(font);
		cell.setCellStyle(style);

		//output file
		try (FileOutputStream out = new FileOutputStream(BOOK_PATH + BOOK_NAME);) {
			book.write(out);
		} catch (IOException e) {
			// TODO: handle exception
			System.err.println(e.getStackTrace());
		}
	}

	private void executeUpdateBook() {

		try (FileInputStream in = new FileInputStream(BOOK_PATH + BOOK_NAME); ) {

			try (Workbook book = WorkbookFactory.create(in); ) {

				Sheet sheet = book.getSheet(SHEET_NAME);
				Row row = sheet.getRow(0);
				Cell cell = row.getCell(0);

				cell.setCellValue("NICE TO MEET YOU!");

				try (FileOutputStream out = new FileOutputStream(BOOK_PATH + BOOK_NAME); ) {
					book.write(out);

				} catch (IOException e) {
					// TODO: handle exception
					System.err.println(e.getStackTrace());
				}

			} catch (InvalidFormatException e) {
				// TODO: handle exception
				System.err.println(e.getStackTrace());
			}

		} catch (IOException e) {
			// TODO: handle exception
			System.err.println(e.getStackTrace());
		}


	}

}
