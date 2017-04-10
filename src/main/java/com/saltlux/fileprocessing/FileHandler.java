package com.saltlux.fileprocessing;

import java.util.List;

import com.google.gson.JsonObject;

public abstract class FileHandler {
	protected String DATE_FORMAT = "yyyyMMdd HH:mm:ss";
	
	protected int rowsCount = 0;
	protected boolean firstRowHeader = true;
	protected int skipFirstRows = 0;
	protected int skipLastRows = 0;
	protected String fileName = null;
	protected String filePath = null;
	protected boolean addTimestampColumn = true;
	
	public FileHandler(String fileName, String filePath) {
		this.fileName = fileName;
		this.filePath = filePath;
	}
	
	public FileHandler(String fileName, String filePath, 
			boolean firstRowHeader, int skipFirstRows, int skipLastRows) {
		this(fileName, filePath);
		this.firstRowHeader = firstRowHeader;
		this.skipFirstRows = skipFirstRows;
		this.skipLastRows = skipLastRows;
	}
	
	/**
	 * Get list contains all sheets in excel file or name for the delimiter file
	 * 
	 * @return
	 */
	abstract List<String> getSheetList();
	
	/**
	 * Get list of field name in table
	 * 
	 * @return
	 */
	abstract List<String> getFieldList();
	
	
	/**
	 * Get one data page from data source
	 * 
	 * @param sheetName:
	 *            name of the sheet to get data (excel case)
	 * @param page:
	 *            zero based page number
	 * @param pageSize:
	 *            number of row in 1 page
	 * @return a JsonObject contains: total row, column names, data rows
	 */
	abstract JsonObject getDataPage(String sheetName, int page, int pageSize);		
	
	/**
	 * Get all data from source
	 * 
	 * @param sheetName:
	 *            name of the sheet to get data (excel case)
	 * @param colList:
	 *            get only columns in this list or all if null
	 * @return
	 */
	abstract JsonObject getAllData(String sheetName, List<String> colList);
	
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
