package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.google.gson.JsonObject;

public abstract class DefaultHandler {
	protected String DATE_FORMAT = "yyyyMMdd HH:mm:ss";
	
	protected int rowsCount = 0;
	protected boolean firstRowHeader = true;
	protected int skipFirstRows = 0;
	protected int skipLastRows = 0;
	protected String filePath = null;
	protected boolean addTimestampColumn = true;
	
	private File originalFile;
	
	public DefaultHandler(String filePath) throws FileNotFoundException {
		this.filePath = filePath;
		this.originalFile = new File(filePath);
		
		// Throw exception if could not find file
		if (!originalFile.exists()) {
			throw new FileNotFoundException("Not found or not a file: " + originalFile.getPath());
		}
	}
	
	public DefaultHandler(String filePath, boolean firstRowHeader, int skipFirstRows, int skipLastRows) throws FileNotFoundException {
		this(filePath);
		this.firstRowHeader = firstRowHeader;
		this.skipFirstRows = skipFirstRows;
		this.skipLastRows = skipLastRows;
	}
	
	public File getOriginalFile() {
		return originalFile;
	}
	
	public String getFileName() {
		return originalFile.getName();
	}
	
	protected boolean isNotEmptySheetName(String sheetName) {
		return StringUtils.isNotEmpty(sheetName);
	}
	
	/**
	 * Get list contains all sheets in excel file or name for the delimiter file
	 * 
	 * @return
	 */
	public abstract List<String> getSheetNames();
	
	/**
	 * Get all data from source
	 * 
	 * @param sheetName:
	 *            Name of the sheet to get data (excel case)
	 * @return
	 */
	public abstract JsonObject getSheetData(String sheetName);
	
	public abstract JsonObject getSheetData(String sheetName, int maxRows);
	
	public abstract JsonObject getFirstSheetData();
	
	public abstract JsonObject getFirstSheetData(int maxRows);
}
