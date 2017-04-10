package com.saltlux.fileprocessing;

import java.io.File;
import java.util.List;

import org.junit.Test;

import com.incesoft.tools.excel.xlsx.Cell;
import com.incesoft.tools.excel.xlsx.Sheet;
import com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;

public class SJXLSXTest {

	@Test
	public void testPerformance01() {
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		long start = System.currentTimeMillis();
		SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(new File(filePath));
		
		// medium data set,just load all at a time
        Sheet sheetToRead = workbook.getSheet(0);
        List<Cell[]> rows = sheetToRead.getRows();
        int rowPos = 0;
        for (Cell[] row : rows) {
                rowPos++;
        }
        
        System.out.println("Read " + rowPos + " rows in " + (System.currentTimeMillis() - start) + "ms");
	}
	
	@Test
	public void testPerformance02() {
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		long start = System.currentTimeMillis();
		SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(new File(filePath));
		
		// here we assume that the sheet contains too many rows which will leads
        // to memory overflow;
        // So we get sheet without loading all records
        Sheet sheetToRead = workbook.getSheet(0, false);
        SheetRowReader reader = sheetToRead.newReader();
        Cell[] row;
        int rowPos = 0;
        while ((row = reader.readRow()) != null) {
                rowPos++;
        }
        
        System.out.println("Read " + rowPos + " rows in " + (System.currentTimeMillis() - start) + "ms");
	}

}
