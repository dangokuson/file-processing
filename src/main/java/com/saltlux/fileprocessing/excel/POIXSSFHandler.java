package com.saltlux.fileprocessing.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.monitorjbl.xlsx.StreamingReader;
import com.saltlux.fileprocessing.DefaultHandler;
import com.saltlux.fileprocessing.Response;

public class POIXSSFHandler extends DefaultHandler {
	
	public static final String XSSF_EXTENSION = "XLSX";
	
	public POIXSSFHandler(String filePath) throws FileNotFoundException, FileFormatException {
		super(filePath);
		
		if (isNotXSSFFile(filePath)) {
			throw new FileFormatException("Excel cannot open " + filePath + 
					" because the file extension is not valid. Only supports XLSX.");
		}
	}

	public POIXSSFHandler(String filePath, boolean firstRowHeader, int skipFirstRows, int skipLastRows) 
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
		
		// Please don't change code format here!
		// FIXME: Use Closable to close resource stream
		try (
			final InputStream is = new FileInputStream(getOriginalFile());
			final Workbook wb = StreamingReader.builder()
					.rowCacheSize(100)		// number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)		// buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is)) {			// InputStream or File for XLSX file (required)
			int numberOfSheets = wb.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++) {
				sheetNames.add(wb.getSheetName(i));
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return sheetNames;
	}
	
	@Override
	public JsonObject getFirstSheetData() {
		return getSheetData(null, -1);
	}
	
	@Override
	public JsonObject getFirstSheetData(int maxRows) {
		return getSheetData(null, maxRows);
	}
	
	@Override
	public JsonObject getSheetData(String sheetName) {
		return getSheetData(sheetName, -1);
	}
	
	@Override
	public JsonObject getSheetData(String sheetName, int maxRows) {
		JsonObject sheetContent = new JsonObject();
		
		// Please don't change code format here!
		// FIXME: Use Closable to close resource stream
		try (
			final InputStream is = new FileInputStream(getOriginalFile());
			final Workbook wb = StreamingReader.builder()
					.rowCacheSize(100)		// number of rows to keep in memory (defaults to 10)
					.bufferSize(4069)		// buffer size to use when reading InputStream to file (defaults to 1024)
					.open(is)) {			// InputStream or File for XLSX file (required)
			Sheet sheet = isNotEmptySheetName(sheetName) ? wb.getSheet(sheetName) : wb.getSheetAt(0);
			sheetContent = parseSheet(sheet, maxRows);
		} catch (IOException ex) {
			ex.printStackTrace();
		}
		return sheetContent;
	}
	
	private boolean isXSSFFile(String filePath) {
		return XSSF_EXTENSION.equalsIgnoreCase(FilenameUtils.getExtension(filePath));
	}
	
	private boolean isNotXSSFFile(String filePath) {
		return !isXSSFFile(filePath);
	}
	
	private JsonObject parseSheet(Sheet sheet, int maxRows) {
		if (sheet == null) {
			return new Response.Failure("There is not sheet is match with given name.").getResponse();
		}
		
		// Default is load all rows in the sheet
		if (maxRows < -1) maxRows = -1;

		try {
			JsonObject sheetData = new JsonObject();
			JsonArray dataOfRows = new JsonArray();
			for (Row row : sheet) {
				// Handler for preview rows
				if (maxRows > 0 && row.getRowNum() <= maxRows) break;

				// Parse row data
				dataOfRows.add(parseRow(row));
			}

			// TODO: Adding more information here!
			sheetData.add(Response.DATASET_PROPERTY, dataOfRows);
			
			// Sort the list column indexes
			sortColumnIndexes();
			
			sheetData.addProperty(Response.MAX_COLUMN_PROPERTY, getMaxCellNumber());
			sheetData.add(Response.COLUMNS_PROPERTY, new Gson().toJsonTree(getColumnIndexes(), List.class));

			// Return JSON Object
			// {status: SUCCESS/ERROR, data: {}, error: { code: 000, message: "" }}
			return new Response.Success(sheetData).getResponse();		
		} catch (Exception ex) {
			return new Response.Failure(ex.getLocalizedMessage()).getResponse();
		}
	}
	
	private JsonObject parseRow(Row row) {
		JsonObject dataOfRow = new JsonObject();
		if (row == null) return dataOfRow;

		setMaxCellNumber(row);

		for (Cell cell : row) {
			String columnIdx = String.valueOf(cell.getColumnIndex());
			// Add column's index to the list if it was not exist
			addColumnIndexToList(columnIdx);
			
			dataOfRow.addProperty(columnIdx, cell.getStringCellValue());
		}
		return dataOfRow;
	}
	
}
