package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.monitorjbl.xlsx.StreamingReader;

public class POIXSSFtreamerTest {

	@Test
	public void testPerformance01() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance02() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		long start = System.currentTimeMillis();
		try {
			File file = new File(filePath);
			Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(file);
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				for (Cell c : row) {
					System.out.println(c.getColumnIndex() + ": " + c.getStringCellValue());
				}
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}
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
	
	@Test
	public void testPerformance05() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test08.XLSX").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance06() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test06.xlsx").getPath();
		excelStreamingReader(filePath);
	}
	
	@Test
	public void testPerformance07() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test06.xlsx").getPath();
		long start = System.currentTimeMillis();
		JsonObject objects = new JsonObject();
		try {
			File file = new File(filePath);
			Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(file);
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				for (Cell c : row) {
					objects.addProperty(String.valueOf(c.getColumnIndex()), c.getStringCellValue());
				}
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testPerformance08() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("SampleSum.xlsx").getPath();
		long start = System.currentTimeMillis();
		try {
			File file = new File(filePath);
			Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(file);
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				for (Cell c : row) {
					System.out.println(c.getStringCellValue());
				}
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testPerformance09() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test04.XLSX").getPath();
		long start = System.currentTimeMillis();
		try {
			File file = new File(filePath);
			Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(file);
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				for (Cell c : row) {
					System.out.println(c.getStringCellValue());
				}
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void testPerformance10() throws FileNotFoundException, IOException {
		String filePath = getClass().getClassLoader().getResource("test03.XLSX").getPath();
		long start = System.currentTimeMillis();
		JsonArray datasets = new JsonArray();
		try {
			File file = new File(filePath);
			Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(file);
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				JsonObject object = new JsonObject();
				for (Cell c : row) {
					object.addProperty(String.valueOf(c.getColumnIndex()), c.getStringCellValue());
				}
				datasets.add(object);
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println(datasets.toString());
	}
	
	private void excelStreamingReader(String filePath) throws FileNotFoundException, IOException {
		long start = System.currentTimeMillis();
		JsonObject objects = new JsonObject();
		try (InputStream is = new FileInputStream(new File(filePath));
				Workbook wb = StreamingReader.builder().sstCacheSize(100).bufferSize(4096).open(is);) {
			Sheet sheet = wb.getSheetAt(0);
			long count = 0;
			for (Row row : sheet) {
				for (Cell c : row) {
					objects.addProperty(String.valueOf(c.getColumnIndex()), c.getStringCellValue());
				}
				count++;
			}
			System.out.println("Read " + count + " rows in " + (System.currentTimeMillis() - start) + "ms");
		}
	}

}
