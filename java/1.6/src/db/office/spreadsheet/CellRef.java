package db.office.spreadsheet;

public class CellRef {

	private static int _INT_26_1 = 26;
	private static int _INT_26_2 = 676;
	private static int COLREF_MAX_SIZE = 3;
	
	private int rowIndex;
	private String colRef;
	private int colIndex;

	public static String toBase26(int n) {	
		char[] chars = new char[COLREF_MAX_SIZE];
		int i = chars.length;	
		while (n != 0) {
			chars[--i] = (char)('A' + (n-1)%26);
			n = (n-1)/26;
		}
		return new String(chars, i, COLREF_MAX_SIZE-i);
	}

	public static int fromBase26(String str) {		
		char[] chars = str.toCharArray();
		if (chars.length==1) {
			return (int)chars[0] - 64;
		} else if (chars.length==2) {
			return (int)chars[1] - 64 + ((int)chars[0] - 64)*_INT_26_1;
		} else if (chars.length==3) {
			return (int)chars[2] - 64 + ((int)chars[1] - 64)*_INT_26_1 + ((int)chars[0] - 64)*_INT_26_2;
		}	
		throw new IllegalArgumentException("Input string is too long");
	}
	
	public CellRef(String cellRef) {
		int i = 0;
		while (i < cellRef.length() && !Character.isDigit(cellRef.charAt(i))) {
			i++;
		}
		rowIndex = Integer.parseInt(cellRef.substring(i));
		colRef = cellRef.substring(0, i);
	}
	
	public CellRef(int rowIndex, int colIndex) {
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		colRef = toBase26(colIndex);
	}
	
	public int getRowIndex() {
		return rowIndex;
	}
	
	public String getColRef() {
		return colRef;
	}
	
	public int getColIndex() {
		return colIndex;
	}
	
	public String toString() {
		return colRef + String.valueOf(rowIndex);
	}
	
	public CellRef copyToRow(int rowIndex) {
		return new CellRef(rowIndex, colIndex);
	}
	
	public CellRef copyToColumn(int colIndex) {
		return new CellRef(rowIndex, colIndex);
	}
	
}
