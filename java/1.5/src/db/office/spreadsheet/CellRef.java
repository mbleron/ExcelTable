package db.office.spreadsheet;

public class CellRef {

	public int row;
	public String column;
	
	public CellRef(String cellRef) {
		int i = 0;
		while (i < cellRef.length() && !Character.isDigit(cellRef.charAt(i))) {
			i++;
		}
		this.row = Integer.parseInt(cellRef.substring(i));
		this.column = cellRef.substring(0, i);
	}
	
}
