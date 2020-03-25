package db.office.spreadsheet.oox;

import java.util.HashMap;
import java.util.Map;
import db.office.spreadsheet.ICellType;

public enum OOXCellType implements ICellType {

	SHAREDSTRING("s"),
	NUMBER("n"),
	BOOLEAN("b"),
	DATETIME("d"),
	INLINESTR("inlineStr"),
	ERROR("e"),
	STRING("str"),
	UNKNOWN("");

	private static final Map<String,OOXCellType> CT_CACHE = new HashMap<String,OOXCellType>(5,1);
	
	static {
		for (OOXCellType vt:values()) {
			CT_CACHE.put(vt.label, vt);
		}
	}
	
	private final String label;
	
	private OOXCellType(String label) {
		this.label = label;
	}
	
	public static OOXCellType get(String label) {
		OOXCellType cellType = CT_CACHE.get(label);
		return (cellType != null)?cellType:OOXCellType.UNKNOWN;
	}
	
	public String getLabel() {
		return label;
	}
}
