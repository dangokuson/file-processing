package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.gson.JsonObject;

public class XSSFFileHandler extends FileHandler {
	
	public XSSFFileHandler(String fileName, String filePath) {
		super(fileName, filePath);
	}

	public XSSFFileHandler(String fileName, String filePath, 
			boolean firstRowHeader, int skipFirstRows, int skipLastRows) {
		super(fileName, filePath, firstRowHeader, skipFirstRows, skipLastRows);
	}

	@Override
	public List<String> getSheetList() {
		List<String> sheetNames = new ArrayList<>();
		Workbook wb = null;
		
		try {
			/* http://poi.apache.org/spreadsheet/quick-guide.html#FileInputStream
			 * 
			 * Using a File object allows for lower memory consumption, while an
			 * InputStream requires more memory as it has to buffer the whole
			 * file
			 */
			File file = new File(filePath);
			if (!file.exists()) {
				throw new FileNotFoundException("Could not find file [" + file.getPath() + "]");
			}
			wb = WorkbookFactory.create(file);
			int numberOfSheets = wb.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++) {
				sheetNames.add(wb.getSheetName(i));
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
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
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	JsonObject getDataPage(String sheetName, int page, int pageSize) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	JsonObject getAllData(String sheetName, List<String> colList) {
		// TODO Auto-generated method stub
		return null;
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

}
