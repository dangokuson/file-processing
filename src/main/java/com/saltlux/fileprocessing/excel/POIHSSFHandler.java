package com.saltlux.fileprocessing.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

import com.google.gson.JsonObject;
import com.saltlux.fileprocessing.DefaultHandler;

public class POIHSSFHandler extends DefaultHandler {
	
	public static final String HSSF_EXTENSION = "XLS";

	public POIHSSFHandler(String filePath) throws FileNotFoundException, FileFormatException {
		super(filePath);
		
		if (isNotHSSFFile(filePath)) {
			throw new FileFormatException("Excel cannot open " + filePath + 
					" because the file extension is not valid. Only supports XLS.");
		}
	}
	
	public POIHSSFHandler(String filePath, boolean firstRowHeader, int skipFirstRows, int skipLastRows) 
			throws FileNotFoundException, FileFormatException {
		super(filePath, firstRowHeader, skipFirstRows, skipLastRows);
		
		if (isNotHSSFFile(filePath)) {
			throw new FileFormatException("Excel cannot open " + filePath + 
					" because the file extension is not valid. Only supports XLS.");
		}
	}

	@Override
	public List<String> getSheetNames() {
		final List<String> sheetNames = new ArrayList<>();
		
		try (
			final InputStream is = new FileInputStream(getOriginalFile());
			final HSSFWorkbook wb = new HSSFWorkbook(is)) {

			int numberOfSheets = wb.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++) {
				sheetNames.add(wb.getSheetName(i));
			}
		} catch (IOException ex) {
			ex.printStackTrace();
		}
		return sheetNames;
	}

	@Override
	public JsonObject getSheetData(String sheetName) {
		return getSheetData(sheetName, -1);
	}

	@Override
	public JsonObject getSheetData(String sheetName, int maxRows) {
		return null;
	}

	@Override
	public JsonObject getFirstSheetData() {
		return getSheetData(null, -1);
	}

	@Override
	public JsonObject getFirstSheetData(int maxRows) {
		return getSheetData(null, maxRows);
	}
	
	private boolean isHSSFFile(String filePath) {
		return HSSF_EXTENSION.equalsIgnoreCase(FilenameUtils.getExtension(filePath));
	}
	
	private boolean isNotHSSFFile(String filePath) {
		return !isHSSFFile(filePath);
	}

}
