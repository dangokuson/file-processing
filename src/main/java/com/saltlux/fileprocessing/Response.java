package com.saltlux.fileprocessing;

import com.google.gson.JsonObject;

public class Response {
	static final String STATUS_PROPERTY = "status";
	static final String DATA_PROPERTY = "data";
	
	static final String ERROR_PROPERTY = "error";
	
	static final String ERROR_CODE_PROPERTY = "errorCode";
	static final String ERROR_MESSAGE_PROPERTY = "message";
	
	public static final String DATASET_PROPERTY = "dataset";
	public static final String COLUMNS_PROPERTY = "columns";
	public static final String MAX_COLUMN_PROPERTY = "maxcol";
	

	public static class Failure {
		private static int errorCode;
		private static String message;
		
		public Failure(int errorCode, String errorMsg) {
			Failure.errorCode = errorCode;
			Failure.message = errorMsg;
		}
		
		public Failure(String errorMsg) {
			// Default error code is 8000 (8000, 8001, 8002, ....)
			Failure.errorCode = 8000; // Common error
			Failure.message = errorMsg;
		}

		public JsonObject getResponse() {
			JsonObject resp = new JsonObject();
			resp.addProperty(STATUS_PROPERTY, ResponseType.Failure.name());
			resp.add(DATA_PROPERTY, new JsonObject());
			resp.add(ERROR_PROPERTY, createErrorObject(errorCode, message));
			return resp;
		}
		
		private static JsonObject createErrorObject(int errorCode, String errorMsg) {
			JsonObject error = new JsonObject();
			error.addProperty(ERROR_CODE_PROPERTY, errorCode);
			error.addProperty(ERROR_MESSAGE_PROPERTY, errorMsg);
			return error;
		}
	}
	

	public static class Success {
		private static JsonObject jsonObject;
		
		public Success(JsonObject jsonObject) {
			Success.jsonObject = jsonObject;
		}
		
		public JsonObject getResponse() {
			JsonObject resp = new JsonObject();
			resp.addProperty(STATUS_PROPERTY, ResponseType.Success.name());
			resp.add(DATA_PROPERTY, jsonObject);
			resp.add(ERROR_PROPERTY, new JsonObject());
			return resp;
		}
	}
}
