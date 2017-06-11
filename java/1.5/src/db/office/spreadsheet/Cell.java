package db.office.spreadsheet;

import java.sql.Clob;
import java.sql.Connection;
import java.sql.SQLException;

import oracle.sql.ANYDATA;
import oracle.sql.CHAR;
import oracle.sql.CLOB;
import oracle.sql.Datum;

public class Cell {

	//private static CharacterSet charset = CharacterSet.make(CharacterSet.DEFAULT_CHARSET);
	
	private CellRef ref;
	private String type;
	private String value;
	
	public Cell (CellRef cellRef, String value, String type) {
		this.ref = cellRef;
		this.type = type;
		this.value = value;
	}
	
	public String toString() {
		return this.ref.column + this.ref.row + ":" + this.value;
	}
	
	public Object[] getOraData (Connection conn) throws SQLException {
		ANYDATA data = null;
		if (this.value.length() <= ReadContext.VC2_MAXSIZE) {
			data = ANYDATA.convertDatum(new CHAR(this.value, null));
		} else {
			//Clob lobdata = conn.createClob();
			Clob lobdata = CLOB.createTemporary(conn, false, CLOB.DURATION_SESSION);
			lobdata.setString(1, this.value);
			data = ANYDATA.convertDatum((Datum) lobdata);
		}
		return new Object[] {this.ref.row, this.ref.column, this.type, data};
	}
	
}
