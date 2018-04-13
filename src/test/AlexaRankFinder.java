package test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class AlexaRankFinder {
	private static String EXCELPATH = "D:\\ExcelSheet.xlsx";
	private static String DRIVERPATH = "D:\\No Longer Using\\Softwares\\drivers\\chromedriver.exe";
	
	private static String URL = "https://www.alexa.com/siteinfo/tamiltechies.in";
	public FileInputStream fis = null;
	public FileOutputStream fos = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	String xlFilePath;
	static LocalDate localDate = LocalDate.now();
	static String today = DateTimeFormatter.ofPattern("dd/MM/yyy").format(localDate);
	
	public AlexaRankFinder(String xlFilePath) throws Exception {
		this.xlFilePath = xlFilePath;
		fis = new FileInputStream(xlFilePath);
		workbook = new XSSFWorkbook(fis);
		fis.close();
	}

	public boolean setCellData(String sheetName, int colNumber, int rowNum, String value) {
		try {
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);

			cell = row.getCell(colNumber);
			if (cell == null)
				cell = row.createCell(colNumber);

			cell.setCellValue(value);

			fos = new FileOutputStream(xlFilePath);
			workbook.write(fos);
			fos.close();
			System.out.println("Entry added for: "+today);
		} catch (Exception ex) {
			ex.printStackTrace();
			return false;
		}
		return true;
	}

	private static List<String> getRanksAndDate() {

		System.setProperty("webdriver.chrome.driver", DRIVERPATH);
		WebDriver driver = new ChromeDriver();
		driver.get(URL);
		WebElement global = driver.findElement(
				By.cssSelector(".globleRank > span:nth-child(1) > div:nth-child(2) > strong:nth-child(2)"));
		WebElement search = driver.findElement(
				By.cssSelector(".countryRank > span:nth-child(1) > div:nth-child(2) > strong:nth-child(2)"));
		List<String> values = new ArrayList<>();
		values.add(today);
		values.add(search.getText());
		values.add(global.getText());
		driver.close();
		return values;

	}

	private static int getNumberOfRowsInExcel() {
		try {
			InputStream is = new FileInputStream(EXCELPATH);
			Workbook wb = WorkbookFactory.create(is);
			Sheet sheet = wb.getSheetAt(0);
			Iterator rowIter = sheet.rowIterator();
			Row r = (Row) rowIter.next();
			short lastCellNum = r.getLastCellNum();
			int[] dataCount = new int[lastCellNum];
			int col = 0;
			rowIter = sheet.rowIterator();
			while (rowIter.hasNext()) {
				Iterator cellIter = ((Row) rowIter.next()).cellIterator();
				while (cellIter.hasNext()) {
					Cell cell = (Cell) cellIter.next();
					col = cell.getColumnIndex();
					dataCount[col] += 1;
				}
			}
			is.close();
			if (wb.getSheetAt(0).getRow(dataCount[col] - 1).getCell(0).toString().equals(today)) {
				System.err.println("Already Rank checked");
				return 0;
			}
			return dataCount[col];
		} catch (Exception e) {
			e.printStackTrace();
		}
		return 0;
	}

	public static void main(String args[]) throws Exception {
		AlexaRankFinder ems = new AlexaRankFinder(EXCELPATH);
		int row = 0;
		int column = getNumberOfRowsInExcel();
		if (column != 0) {
			List<String> rankAndDate = getRanksAndDate();
			Iterator<String> iter = rankAndDate.iterator();
			while (iter.hasNext()) {
				for (int i = 0; i < rankAndDate.size(); i++) {
					row = i;
					ems.setCellData("Rank Checker", row, column, iter.next());
				}
			}
		}
	}
}
