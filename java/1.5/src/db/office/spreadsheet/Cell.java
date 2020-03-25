package db.office.spreadsheet;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.Clob;
import java.sql.Connection;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import oracle.sql.ANYDATA;
import oracle.sql.BINARY_DOUBLE;
import oracle.sql.CHAR;
import oracle.sql.CLOB;
import oracle.sql.Datum;
import oracle.sql.TIMESTAMP;

public class Cell {

	private static final DateFormat DATE_FMT = new SimpleDateFormat("yyyy-MM-dd");
	private static final DateFormat DATETIME_FMT = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
	
	public static final String CT_SHAREDSTRING = "s";
	public static final String CT_NUMBER = "n";
	public static final String CT_BOOLEAN = "b";
	public static final String CT_DATETIME = "d";
	public static final String CT_INLINESTR = "inlineStr";
	private static final String BOOL_TRUE = "TRUE";
	private static final String BOOL_FALSE = "FALSE";
	
	private CellRef ref;
	private String type;
	private String value;
	private int sheetIndex;
	
	public Cell (CellRef cellRef, String value, String type, int sheetIndex) {
		this.ref = cellRef;
		this.type = type;
		this.value = value;
		this.sheetIndex = sheetIndex;
	}
	
	public String toString() {
		return this.ref.column + this.ref.row + ":" + this.value;
	}

	public String getValue() {
		return this.value;
	}
	
	private TIMESTAMP getTimestamp() throws ParseException {
		
		long time;
		int sepIndex;
		
		if (this.value.length() == 10) {
			time = DATE_FMT.parse(this.value).getTime();
		} else if ((sepIndex = this.value.lastIndexOf('.')) != -1) {
			BigDecimal fsecs = new BigDecimal(this.value.substring(sepIndex)).setScale(3, RoundingMode.HALF_UP);
			time = DATETIME_FMT.parse(this.value).getTime() + (long)(fsecs.doubleValue()*1000);
		} else {
			time = DATETIME_FMT.parse(this.value).getTime();
		}
		
		return new TIMESTAMP(new java.sql.Timestamp(time));

	}
	
	public Object[] getOraData (Connection conn) throws CellReaderException {
		try {
			ANYDATA data = null;
			if (this.type == null || CT_NUMBER.equals(this.type)) {
				BINARY_DOUBLE bdouble = null;
				if (this.value != null && this.value.length() != 0) {
					//contrary to what the documentation says, BINARY_DOUBLE(String) is not implemented on 11.2.0.1
					bdouble = new BINARY_DOUBLE(Double.parseDouble(this.value));
				}
				data = ANYDATA.convertDatum(bdouble);

			} else if (CT_BOOLEAN.equals(this.type)) {
				String boolString = ("1".equals(this.value))?BOOL_TRUE:BOOL_FALSE;
				data = ANYDATA.convertDatum(new CHAR(boolString, null));

			} else if (CT_DATETIME.equals(this.type)) {
				data = ANYDATA.convertDatum(this.getTimestamp());

			} else {

				CHAR chars = new CHAR(this.value, null);
				if (chars.getLength() <= ReadContext.VC2_MAXSIZE) {
					data = ANYDATA.convertDatum(chars);
				} else {
					Clob lobdata = CLOB.createTemporary(conn, false, CLOB.DURATION_SESSION);
					lobdata.setString(1, this.value);
					data = ANYDATA.convertDatum((Datum) lobdata);
				}

			}

			return new Object[] {this.ref.row, this.ref.column, this.type, data, this.sheetIndex};

		} catch (Exception e) {
			throw new CellReaderException("Error in getOraData() method", e);
		}

	}
	
}
