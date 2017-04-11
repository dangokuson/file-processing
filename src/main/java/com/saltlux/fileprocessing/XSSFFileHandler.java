package com.saltlux.fileprocessing;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

import com.google.gson.JsonObject;
import com.monitorjbl.xlsx.StreamingReader;

public class XSSFFileHandler extends FileHandler {
	
	public static final String XSSF_EXTENSION = "XLSX";
	
	public XSSFFileHandler(String filePath) throws FileNotFoundException, FileFormatException {
		super(filePath);
		
		if (isNotXSSFFile(filePath)) {
			throw new FileFormatException("Excel cannot open " + filePath + 
					" because the file extension is not valid. Only supports XLSX.");
		}
	}

	public XSSFFileHandler(String filePath, boolean firstRowHeader, int skipFirstRows, int skipLastRows) 
			throws FileNotFoundException, FileFormatException {
		super(filePath, firstRowHeader, skipFirstRows, skipLastRows);
		
		if (isNotXSSFFile(filePath)) {
			throw new FileFormatException("Excel cannot open " + filePath + 
					" because the file extension is not valid. Only supports XLSX.");
		}
	}

	@Override
	public List<String> getSheetNames() {
		final List<String> sheetNames = new ArrayList<>();
		InputStream is = null;
		Workbook wb = null;
		
		try {
			is = new FileInputStream(getOriginalFile());
			wb = StreamingReader.builder()
					.rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is);            // InputStream or File for XLSX file (required)
			int numberOfSheets = wb.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++) {
				sheetNames.add(wb.getSheetName(i));
			}		
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException ex) {
					ex.printStackTrace();
				}
			}
			
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException ex) {
					ex.printStackTrace();
				}
			}
		}
		return sheetNames;
	}

	@Override
	List<String> getFieldList() {
		return null;
	}

	@Override
	public JsonObject getAllData(String sheetName) {
		JsonObject sheetContent = new JsonObject();

		InputStream is = null;
		Workbook wb = null;
		
		try {
			is = new FileInputStream(getOriginalFile());
			wb = StreamingReader.builder()
					.rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is);            // InputStream or File for XLSX file (required)
			Sheet sheet = wb.getSheet(sheetName);
			sheetContent = parseSheet(sheet);
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException ex) {
					ex.printStackTrace();
				}
			}
			
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException ex) {
					ex.printStackTrace();
				}
			}
		}
		return sheetContent;
	}

	@Override
	boolean saveImportFile(String desFile, String selectedSheet) {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	boolean updateDataFile(JsonObject rowData, boolean isSwitch) {
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	JsonObject getAllData(String sheetName, List<String> colList, String columnOption) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	JsonObject getFileData(String sheetName, int offset, int limit, List<String> colList, String columnOption) {
		// TODO Auto-generated method stub
		return null;
	}

	private boolean isXSSFFile(String filePath) {
		return XSSF_EXTENSION.equalsIgnoreCase(FilenameUtils.getExtension(filePath));
	}
	
	private boolean isNotXSSFFile(String filePath) {
		return !isXSSFFile(filePath);
	}
	
	/**
	 * Parses and shows the content of one sheet
	 * 
	 * @param sheet
	 * @return
	 */
	private JsonObject parseSheet(Sheet sheet) {
		if (sheet == null) {
			throw new IllegalArgumentException("Sheet name is invalid.");
		}
		
		// Return JSON Object
		// {status: SUCCESS/ERROR, data: [], error: { code: 000, message: "" }}
		return null;
	}
}
