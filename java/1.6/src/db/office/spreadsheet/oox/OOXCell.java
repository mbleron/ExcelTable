package db.office.spreadsheet.oox;

import java.sql.Clob;
import java.sql.Connection;

import db.office.spreadsheet.Cell;
import db.office.spreadsheet.CellReaderException;
import db.office.spreadsheet.CellRef;
import db.office.spreadsheet.ReadContext;
import oracle.sql.ANYDATA;
import oracle.sql.BINARY_DOUBLE;
import oracle.sql.CHAR;
import oracle.sql.Datum;

public class OOXCell extends Cell<OOXCellType> {
	
	private static final String BOOL_TRUE = "TRUE";
	private static final String BOOL_FALSE = "FALSE";
	
	public OOXCell(CellRef cellRef, String value, OOXCellType type, int sheetIndex) {
		super(cellRef, value, type, sheetIndex);
	}
	
	public Object[] getOraData (Connection conn) throws CellReaderException {
		try {
			ANYDATA data = null;
			switch (cellType) {
			case NUMBER:
				BINARY_DOUBLE bdouble = null;
				if (this.value != null && !this.value.isEmpty()) {
					bdouble = new BINARY_DOUBLE(this.value);
				}
				data = ANYDATA.convertDatum(bdouble);
				break;

			case SHAREDSTRING:
			case INLINESTR:
			case STRING:
			case ERROR:
				CHAR chars = new CHAR(this.value, null);
				if (chars.getLength() <= ReadContext.VC2_MAXSIZE) {
					data = ANYDATA.convertDatum(chars);
				} else {
					Clob lobdata = conn.createClob();
					lobdata.setString(1, this.value);
					data = ANYDATA.convertDatum((Datum) lobdata);
				}
				break;
				
			case BOOLEAN:
				String boolString = ("1".equals(this.value))?BOOL_TRUE:BOOL_FALSE;
				data = ANYDATA.convertDatum(new CHAR(boolString, null));
				break;
				
			case DATETIME:
				data = ANYDATA.convertDatum(this.getTimestamp());
				break;
				
			default:
				data = ANYDATA.convertDatum(new CHAR(value, null));
				
			}
			
			return new Object[] {this.ref.getRowIndex(), this.ref.getColRef(), this.cellType.getLabel(), data, this.sheetIndex, this.getAnnotation()};
			
		} catch (Exception e) {
			throw new CellReaderException("Error in getOraData() method", e);
		}
	}

	public Cell<?> copyToRow(int rowIndex) {
		throw new java.lang.UnsupportedOperationException("Not supported for OOX Cell");
	}

	public Cell<?> copyToColumn(int colIndex) {
		throw new java.lang.UnsupportedOperationException("Not supported for OOX Cell");
	}

}
