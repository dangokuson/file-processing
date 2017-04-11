package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.List;

import com.google.gson.JsonObject;

public abstract class FileHandler {
	protected String DATE_FORMAT = "yyyyMMdd HH:mm:ss";
	
	protected int rowsCount = 0;
	protected boolean firstRowHeader = true;
	protected int skipFirstRows = 0;
	protected int skipLastRows = 0;
	protected String filePath = null;
	protected boolean addTimestampColumn = true;
	
	private File originalFile;
	
	public FileHandler(String filePath) throws FileNotFoundException {
		this.filePath = filePath;
		this.originalFile = new File(filePath);
		
		// Throw exception if could not find file
		if (!originalFile.exists()) {
			throw new FileNotFoundException("Not found or not a file: " + originalFile.getPath());
		}
	}
	
	public FileHandler(String filePath, boolean firstRowHeader, int skipFirstRows, int skipLastRows) throws FileNotFoundException {
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
	
	/**
	 * Get list contains all sheets in excel file or name for the delimiter file
	 * 
	 * @return
	 */
	abstract List<String> getSheetNames();
	
	/**
	 * Get list of field name in table
	 * 
	 * @return
	 */
	abstract List<String> getFieldList();
	
	/**
	 * Get all data from source
	 * 
	 * @param sheetName:
	 *            Name of the sheet to get data (excel case)
	 * @return
	 */
	abstract JsonObject getAllData(String sheetName);
	
	/**
	 * Save final selected data (csv file or user selected excel sheet)
	 * 
	 * @return true if success
	 */
	abstract boolean saveImportFile(String desFile, String selectedSheet);
	
	/**
	 * Update data row
	 * 
	 * @param rowData:
	 *            contain data to update {'rowNum':'rowData'}
	 * @isSwitch: switch columns and rows
	 * @return
	 */
	abstract boolean updateDataFile(JsonObject rowData, boolean isSwitch);
	
	abstract JsonObject getAllData(String sheetName, List<String> colList, String columnOption);
	
	abstract JsonObject getFileData(String sheetName, int offset, int limit, List<String> colList, String columnOption);
}
