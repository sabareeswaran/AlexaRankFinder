package test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Test {
	private static String EXCELPATH = "D:\\ExcelSheet.xlsx";

	public static void main(String[] args) {
		
		/*String[] str = {"a", "b","c"};
		List<String> strList = new ArrayList<String>();
		for ( String obj : strList) {
			System.out.println(obj);
		}
		
		for (int i=0; i<strList.size(); i++) {
			System.out.println(strList.get(i));			
		}*/

		
		List<String> al = new ArrayList<String>();
		al.add("a");
		al.add("b");
		al.add("c");
		Iterator<String> iter = al.iterator();
		
		while(iter.hasNext()) {
			
			for(int i=0; i<al.size();i++) {
				System.out.println(iter.next());
			}
		}
		
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
