package db.office.spreadsheet;

import java.util.ArrayList;

public class Row extends ArrayList<Cell<?>> {

	private static final long serialVersionUID = 6615809380592083862L;
	private int rowIndex;
	
	public Row(int rowIndex, int cellCount) {
		super(cellCount);
		this.rowIndex = rowIndex;
	}
	
	public int getRowIndex() {
		return rowIndex;
	}
	
	public Cell<?>[] getCells() {
		Cell<?>[] cells = new Cell[size()];
		return toArray(cells);
	}
	
	public static Row copyTo(Row row, int rowIndex) {
		Row targetRow = null;
		if (row != null) {
			targetRow = new Row(rowIndex, row.size());
			for (Cell<?> c:row) {
				targetRow.add(c.copyToRow(rowIndex));
			}
		}
		return targetRow;		
	}
	
}
