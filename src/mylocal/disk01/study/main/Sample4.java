package mylocal.disk01.study.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

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

	private static final SimpleDateFormat YYYYMM_FORMAT = new SimpleDateFormat("yyyyMM");
	private static final String YYYYMM = YYYYMM_FORMAT.format( new Date() );

	private static final String EXCEL_EXTENSION = ".xlsx";
	private static final String BOOK_NAME = YYYYMM + EXCEL_EXTENSION;

	private static final String BOOK_PATH = "C:\\Dev\\";
	//private static final String BOOK_NAME = "Sample4.xlsx";
	private static final String SHEET_NAME = "mySheet";

	private static final String CURRENT_DIR = System.getProperty("user.dir", BOOK_PATH);


	public static void main(String[] args) {
		// TODO 自動生成されたメソッド・スタブ
		new Sample4().execute();
		new Sample4().executeCreateBook();
//		new Sample4().executeUpdateBook();
//		new Sample4().getCurrentPath();
//		new Sample4().getYYYYMM();
		new Sample4().createExcel();
	}

	private void execute() {
		System.out.println("Sample4 Control Excel");
	}

	private void executeCreateBook() {

		try (Workbook book = new XSSFWorkbook(); ) {

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

	private void getCurrentPath() {
		//Properties props = System.getProperties();
		//props.list(System.out);

		System.getProperties().list(System.out);

		String dir = System.getProperty("user.dir", "c:\\dev");
		System.out.println(dir);
	}

	private void getYYYYMM() {

		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

		System.out.println(date.toString());
		System.out.println(sdf.format(date));
		System.out.println(YYYYMM);
	}

	private void createExcel() {

		System.out.println(CURRENT_DIR);
		String filePath = new File(CURRENT_DIR, BOOK_NAME).getPath();
		System.out.println(filePath);

		System.exit(0);

		try (Workbook book = new XSSFWorkbook(); ) {

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
			try (FileOutputStream out = new FileOutputStream(filePath);) {
				book.write(out);
			} catch (IOException e) {
				// TODO: handle exception
				System.err.println(e.getStackTrace());
			}
		} catch (IOException e) {
			// TODO: handle exception
			System.err.println(e.getStackTrace());
		}
	}

}
