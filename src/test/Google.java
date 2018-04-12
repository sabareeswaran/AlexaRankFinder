package test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Google {
	private static String EXCELPATH = "D:\\\\ExcelSheet.xlsx";
	
	public static void main(String[] args) throws IOException {
	
		LocalDate localDate = LocalDate.now();
		String today = DateTimeFormatter.ofPattern("dd/MM/yyy").format(localDate);
		System.setProperty("webdriver.chrome.driver", "D:\\No Longer Using\\Softwares\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("http://alexa.com/siteinfo/tamiltechies.in");
		String title = driver.getTitle();
		String url = driver.getCurrentUrl();
		System.out.println(title + url);
		WebElement global = driver.findElement(By.cssSelector(
				".globleRank > span:nth-child(1) > div:nth-child(2) > strong:nth-child(2)"));
		WebElement search = driver.findElement(
				By.cssSelector(".countryRank > span:nth-child(1) > div:nth-child(2) > strong:nth-child(2)"));
		System.out.println(search.getText());
		System.out.println(global.getText());
		System.out.println(today);
		driver.close();
		int numberOfRowsInExcel = getNumberOfRowsInExcel();
		
	}
	
	private static int getNumberOfRowsInExcel() {
		try {
	        InputStream is = new FileInputStream(EXCELPATH);
	        Workbook wb = WorkbookFactory.create(is);
	        Sheet sheet = wb.getSheetAt(0);
	        Iterator rowIter = sheet.rowIterator();
	        Row r = (Row)rowIter.next();
	        short lastCellNum = r.getLastCellNum();
	        int[] dataCount = new int[lastCellNum];
	        int col = 0;
	        rowIter = sheet.rowIterator();
	        while(rowIter.hasNext()) {
	            Iterator cellIter = ((Row)rowIter.next()).cellIterator();
	            while(cellIter.hasNext()) {
	                Cell cell = (Cell)cellIter.next();
	                col = cell.getColumnIndex();
	                dataCount[col] += 1;
	                DataFormatter df = new DataFormatter();
	            }
	        }
	        is.close();
	        return dataCount[col];
	    }
	    catch(Exception e) {
	        e.printStackTrace();
	    }
		return 0;
	}
}