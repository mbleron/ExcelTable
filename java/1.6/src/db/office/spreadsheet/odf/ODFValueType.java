package db.office.spreadsheet.odf;

import java.util.HashMap;
import java.util.Map;
import db.office.spreadsheet.ICellType;

public enum ODFValueType implements ICellType {
	
	FLOAT("float"),
	STRING("string"), 
	DATE("date"),
	BOOLEAN("boolean"),
	PERCENTAGE("percentage");
	
	private static final Map<String,ODFValueType> VT_CACHE = new HashMap<String,ODFValueType>(5,1);
	
	static {
		for (ODFValueType vt:values()) {
			VT_CACHE.put(vt.label, vt);
		}
	}
	
	private final String label;
	
	private ODFValueType(String label) {
		this.label = label;
	}
	
	public static ODFValueType get(String label) {
		return VT_CACHE.get(label);
	}
	
	public String getLabel() {
		return label;
	}
	
}
