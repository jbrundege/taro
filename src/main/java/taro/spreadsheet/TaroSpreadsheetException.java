package taro.spreadsheet;

import taro.TaroException;

public class TaroSpreadsheetException extends TaroException {

	public TaroSpreadsheetException() {
	}

	public TaroSpreadsheetException(String message) {
		super(message);
	}

	public TaroSpreadsheetException(String message, Throwable cause) {
		super(message, cause);
	}

	public TaroSpreadsheetException(Throwable cause) {
		super(cause);
	}

	public TaroSpreadsheetException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

}
