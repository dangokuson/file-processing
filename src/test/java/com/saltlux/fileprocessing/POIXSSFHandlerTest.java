package com.saltlux.fileprocessing;

import static org.junit.Assert.*;

import java.io.FileNotFoundException;
import java.util.Arrays;
import java.util.List;

import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.google.gson.JsonObject;
import com.saltlux.fileprocessing.excel.POIXSSFHandler;

public class POIXSSFHandlerTest {

	@Before
	public void setUp() throws Exception {
	}

	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void test01() throws FileNotFoundException, FileFormatException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		POIXSSFHandler fileHandler = new POIXSSFHandler(filePath);
		List<String> sheetNames = fileHandler.getSheetNames();
		List<String> sheetNamesExpected = Arrays.asList("Orders", "Returns", "Users");
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");

		assertFalse(sheetNames.isEmpty());
		assertTrue(sheetNames.equals(sheetNamesExpected));
	}
	
	@Test
	public void test02() throws FileNotFoundException, FileFormatException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		POIXSSFHandler fileHandler = new POIXSSFHandler(filePath);
		List<String> sheetNames = fileHandler.getSheetNames();
		List<String> sheetNamesExpected = Arrays.asList("Smile You Can Read Me !");
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");

		assertFalse(sheetNames.isEmpty());
		assertTrue(sheetNames.equals(sheetNamesExpected));
	}
	
	@Test
	public void test03() throws FileNotFoundException, FileFormatException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		POIXSSFHandler fileHandler = new POIXSSFHandler(filePath);
		List<String> sheetNames = fileHandler.getSheetNames();
		List<String> sheetNamesExpected = Arrays.asList("Smile You Can Read Me !");
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");

		assertFalse(sheetNames.isEmpty());
		assertTrue(sheetNames.equals(sheetNamesExpected));
	}
	
	@Test
	public void test04() throws FileNotFoundException, FileFormatException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		POIXSSFHandler fileHandler = new POIXSSFHandler(filePath);
		JsonObject content = fileHandler.getFirstSheetData();
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");

		assertFalse(content.isJsonNull());
	}
	
	@Test
	public void test05() throws FileNotFoundException, FileFormatException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("test04.XLSX").getPath();
		POIXSSFHandler fileHandler = new POIXSSFHandler(filePath);
		JsonObject content = fileHandler.getFirstSheetData();
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");

		assertFalse(content.isJsonNull());
	}

}
