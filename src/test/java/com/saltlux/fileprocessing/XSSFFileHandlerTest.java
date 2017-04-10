package com.saltlux.fileprocessing;

import static org.junit.Assert.*;

import java.util.Arrays;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class XSSFFileHandlerTest {

	@Before
	public void setUp() throws Exception {
	}

	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void test01() {
		String filePath = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		XSSFFileHandler fileHandler = new XSSFFileHandler(null, filePath);
		List<String> sheetNames = fileHandler.getSheetList();
		List<String> sheetNamesExpected = Arrays.asList("Orders", "Returns", "Users");
		
		assertFalse(sheetNames.isEmpty());
		assertTrue(sheetNames.equals(sheetNamesExpected));
	}
	
	@Test
	public void test02() {
		String filePath = getClass().getClassLoader().getResource("sample.xlsx").getPath();
		XSSFFileHandler fileHandler = new XSSFFileHandler(null, filePath);
		List<String> sheetNames = fileHandler.getSheetList();
		List<String> sheetNamesExpected = Arrays.asList("Smile You Can Read Me !");
		
		assertFalse(sheetNames.isEmpty());
		assertTrue(sheetNames.equals(sheetNamesExpected));
	}

}
