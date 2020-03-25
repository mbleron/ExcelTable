package db.office.spreadsheet.odf;

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

public class ODFCell extends Cell<ODFValueType> {
	
	public ODFCell(CellRef cellRef, String value, ODFValueType type, int sheetIndex) {
		super(cellRef, value, type, sheetIndex);
	}
	
	public Object[] getOraData(Connection conn) throws CellReaderException {
		try {
			ANYDATA data = null;
			switch (cellType) {
			case FLOAT:
			case PERCENTAGE:
				BINARY_DOUBLE bdouble = null;
				if (value != null && !value.isEmpty()) {
					bdouble = new BINARY_DOUBLE(value);
				}
				data = ANYDATA.convertDatum(bdouble);
				break;
			
			case BOOLEAN:
				data = ANYDATA.convertDatum(new CHAR(value.toUpperCase(), null));
				break;
				
			case DATE:
				data = ANYDATA.convertDatum(this.getTimestamp());
				break;
				
			default:
				CHAR chars = new CHAR(value, null);
				if (chars.getLength() <= ReadContext.VC2_MAXSIZE) {
					data = ANYDATA.convertDatum(chars);
				} else {
					Clob lobdata = conn.createClob();
					lobdata.setString(1, value);
					data = ANYDATA.convertDatum((Datum) lobdata);
				}
				
			}
			
			return new Object[] {ref.getRowIndex(), ref.getColRef(), cellType, data, sheetIndex, getAnnotation()};
			
		} catch (Exception e) {
			throw new CellReaderException("Error in getOraData() method", e);
		}
	}
	
	public Cell<?> copyToRow(int rowIndex) {
		return new ODFCell(ref.copyToRow(rowIndex), value, cellType, sheetIndex);
	}
	
	public Cell<?> copyToColumn(int colIndex) {
		return new ODFCell(ref.copyToColumn(colIndex), value, cellType, sheetIndex);
	}

}
