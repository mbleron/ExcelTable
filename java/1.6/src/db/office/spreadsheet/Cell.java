package db.office.spreadsheet;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.Connection;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import oracle.sql.TIMESTAMP;

public abstract class Cell<T extends ICellType> {
	
	protected T cellType;
	
	private static final DateFormat DATE_FMT = new SimpleDateFormat("yyyy-MM-dd");
	private static final DateFormat DATETIME_FMT = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
	
	protected CellRef ref;
	protected String value;
	protected int sheetIndex;
	private String annotation;
	
	public Cell (CellRef cellRef, String value, T cellType, int sheetIndex) {
		this.ref = cellRef;
		this.value = value;
		this.sheetIndex = sheetIndex;
		this.cellType = cellType;
	}
	
	public String toString() {
		return ref.toString() + ":" + value;
	}
	
	public String getValue() {
		return value;
	}
	
	public void setValue(String value) {
		this.value = value;
	}
	
	public CellRef getRef() {
		return ref;
	}
	
	public void setAnnotation(String annotation) {
		this.annotation = annotation;
	}
	
	public String getAnnotation() {
		return annotation;
	}
	
	public boolean hasAnnotation() {
		return (annotation != null && !annotation.isEmpty());
	}
	
	public boolean isOf(T cellType) {
		return (this.cellType == cellType);
	}
	
	public TIMESTAMP getTimestamp() throws ParseException {
		
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
	
	public abstract Object[] getOraData (Connection conn) throws CellReaderException;
	
	public abstract Cell<?> copyToRow(int rowIndex);
	
	public abstract Cell<?> copyToColumn(int colIndex);
	
}
