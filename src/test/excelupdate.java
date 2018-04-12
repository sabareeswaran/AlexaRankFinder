package test;

import static org.hamcrest.CoreMatchers.any;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excelupdate {

	public static void main(String[] args) {
		    try {
		        InputStream is = new FileInputStream("D:\\ExcelSheet.xlsx");
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
		        System.out.println(dataCount[col]);
		        /*for(int x = 0; x < dataCount.length; x++) {
		            System.out.println("col " + x + ": " + dataCount[x]);
		        }*/
		    }
		    catch(Exception e) {
		        e.printStackTrace();
		        return;
		    }
		}
	}