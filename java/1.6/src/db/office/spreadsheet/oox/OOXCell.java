package db.office.spreadsheet.oox;

import java.sql.Clob;
import java.sql.Connection;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.TimeZone;

import db.office.spreadsheet.Cell;
import db.office.spreadsheet.CellReaderException;
import db.office.spreadsheet.CellRef;
import db.office.spreadsheet.ReadContext;
import oracle.sql.ANYDATA;
import oracle.sql.BINARY_DOUBLE;
import oracle.sql.CHAR;
import oracle.sql.DATE;
import oracle.sql.Datum;
import oracle.sql.TIMESTAMP;

public class OOXCell extends Cell<OOXCellType> {
	
	private static final String BOOL_TRUE = "TRUE";
	private static final String BOOL_FALSE = "FALSE";
	private static final Calendar BASE_CALENDAR;
	private static final long BASE_TIME;
	
	static {
		BASE_CALENDAR = new GregorianCalendar(1899,Calendar.DECEMBER,30);
		BASE_CALENDAR.setTimeZone(TimeZone.getTimeZone("UTC"));
		BASE_TIME = BASE_CALENDAR.getTimeInMillis();
	}
	
	private FormatFlags flags;
	
	public static class FormatFlags {
		private boolean isTimestamp;
		public FormatFlags(boolean isTimestamp) {
			this.isTimestamp = isTimestamp;
		}
	}
	
	public OOXCell(CellRef cellRef, String value, OOXCellType type, int sheetIndex, FormatFlags flags) {
		super(cellRef, value, type, sheetIndex);
		this.flags = flags;
	}
	
	public Object[] getOraData (Connection conn) throws CellReaderException {
		try {
			ANYDATA data = null;
			switch (cellType) {
			case NUMBER:
				Datum datum = null;
				if (!value.isEmpty()) {
					if (flags != null) {				
						double d = Double.parseDouble(value);
						if (d < 60) {
							d++;
						}
						long time = BASE_TIME + (long)(d*86400000 + .5);
						if (flags.isTimestamp) {
							datum = new TIMESTAMP(new java.sql.Timestamp(time), BASE_CALENDAR);
						} else {
							// round up to nearest second
							time = (long) (Math.floor((time+500)/1000)*1000);
							datum = new DATE(new java.sql.Timestamp(time), BASE_CALENDAR);
						}
					} else {	
						datum = new BINARY_DOUBLE(value);
					}				
				}
				data = ANYDATA.convertDatum(datum);
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
