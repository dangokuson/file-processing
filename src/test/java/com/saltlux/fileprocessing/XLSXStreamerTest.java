package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.monitorjbl.xlsx.StreamingReader;

public class XLSXStreamerTest {

	@Test
	public void testPerformance01() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance02() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance03() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("Sample-Sales-Data.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance04() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	private void excelStreamingReader(String filePath) throws FileNotFoundException, IOException {
		long start = System.currentTimeMillis();
		try (InputStream is = new FileInputStream(new File(filePath));
				Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(is);) {
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		}
	}

}
