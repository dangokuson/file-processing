package com.saltlux.fileprocessing;

import com.google.gson.JsonObject;

public class JSONResponse {
	static final String ERROR = "ERROR";
	static final String SUCCESS = "SUCCESS";
	
	// Properties
	static final String STATUS_PROPERTY = "status";
	static final String DATA_PROPERTY = "data";
	
	static final String ERROR_PROPERTY = "error";
	
	static final String ERROR_CODE_PROPERTY = "code";
	static final String ERROR_MESSAGE_PROPERTY = "message";
	
	// Default error code is 99999
	static final String DEFAULT_ERROR_CODE_PROPERTY = "99999";
	
	public static final String DATASET_PROPERTY = "dataset";
	public static final String COLUMNS_PROPERTY = "columns";
	public static final String MAX_COLUMN_PROPERTY = "maxcol";
	
	/**
	 * Create new success response
	 * 
	 * @param data
	 * @return
	 */
	public static JsonObject createSuccessResp(JsonObject data) {
		JsonObject resp = new JsonObject();
		resp.addProperty(STATUS_PROPERTY, SUCCESS);
		resp.add(DATA_PROPERTY, data);
		resp.add(ERROR_PROPERTY, new JsonObject());
		return resp;
	}
	
	/**
	 * Create new error response
	 * 
	 * @param errorCode
	 * @param errorMsg
	 * @return
	 */
	public static JsonObject createErrorResp(String errorCode, String errorMsg) {
		JsonObject resp = new JsonObject();
		resp.addProperty(STATUS_PROPERTY, ERROR);
		resp.add(DATA_PROPERTY, new JsonObject());
		resp.add(ERROR_PROPERTY, createErrorObject(errorCode, errorMsg));
		return resp;
	}
	
	/**
	 * Create new error response with default error code
	 * 
	 * @param errorMsg
	 * @return
	 */
	public static JsonObject createErrorResp(String errorMsg) {
		return createErrorObject(DEFAULT_ERROR_CODE_PROPERTY, errorMsg);
	}
	
	private static JsonObject createErrorObject(String errorCode, String errorMsg) {
		JsonObject error = new JsonObject();
		error.addProperty(ERROR_CODE_PROPERTY, errorCode);
		error.addProperty(ERROR_MESSAGE_PROPERTY, errorMsg);
		return error;
	}
}
