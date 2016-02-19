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
import org.apache.poi.ss.util.CellRangeAddress;
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
	private static final String FILEPATH = new File(CURRENT_DIR, BOOK_NAME).getPath();


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

		//System.out.println(CURRENT_DIR);
		//String filePath = new File(CURRENT_DIR, BOOK_NAME).getPath();
		System.out.println(FILEPATH);
		//System.exit(0);

		try (Workbook book = new XSSFWorkbook(); ) {

			//style & font set
			CellStyle style = book.createCellStyle();

			//set border
			style.setBorderTop(CellStyle.BORDER_THIN);
			style.setBorderBottom(CellStyle.BORDER_HAIR);
			style.setBorderLeft(CellStyle.BORDER_MEDIUM);
			style.setBorderRight(CellStyle.BORDER_DOTTED);

			Font font = book.createFont();
			font.setColor(IndexedColors.AQUA.getIndex());

			Row  row    = null;
			Cell cell   = null;
			int  rowNum = 0;
			int  colNum = 0;

			Sheet sheet = null;

			for (int i = 0; i < 3; i++) {
				//sheet create & name set
				sheet = book.createSheet(SHEET_NAME+i);

				//create header line
				rowNum = 0;
				colNum = 0;
				row = sheet.createRow(rowNum);

				cell = row.createCell(colNum);
				cell.setCellValue(rowNum+"-"+colNum);  //set cell value

				cell = row.createCell(++colNum);
				cell.setCellValue(rowNum+"-"+colNum+"longlonglongcell");

				cell = row.createCell(++colNum);
				cell.setCellValue(rowNum+"-"+colNum);

				cell = row.createCell(++colNum);
				cell.setCellValue(rowNum+"-"+colNum);

				cell = row.createCell(++colNum);
				cell.setCellValue(rowNum+"-"+colNum);

				cell = row.createCell(++colNum);
				cell.setCellValue(rowNum+"-"+colNum);

				//fixed
				sheet.createFreezePane(1, 1);

				//auto fillter
				//sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, colNum));

				//auto size col
				for (int j = 0; j < colNum; j++) {
					sheet.autoSizeColumn(j, true);
				}

				//create data row
				for (int j = 0; j < 10; j++) {
					rowNum++;
					colNum = 0;

					row = sheet.createRow(rowNum);

					for (int j2 = 0; j2 < 6; j2++) {
						cell = row.createCell(j2);
						cell.setCellValue(rowNum+"-"+j2);

						cell.setCellStyle(style);
					}
					for (int j2 = 0; j2 < 6; j2++) {
						sheet.autoSizeColumn(j2, true);
					}

				}
				sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 3));


			}



			style.setFont(font);
			cell.setCellStyle(style);

			//output file
			try (FileOutputStream out = new FileOutputStream(FILEPATH);) {
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
