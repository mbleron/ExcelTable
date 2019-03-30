package db.office.spreadsheet;

import java.sql.Clob;
import java.sql.Connection;
import java.sql.SQLException;

import oracle.sql.ANYDATA;
import oracle.sql.BINARY_DOUBLE;
import oracle.sql.CHAR;
import oracle.sql.CLOB;
import oracle.sql.Datum;

public class Cell {
	
	public static final String CT_SHAREDSTRING = "s";
	public static final String CT_NUMBER = "n";
	public static final String CT_BOOLEAN = "b";
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
	
	public Object[] getOraData (Connection conn) throws SQLException {
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
	}
	
}
