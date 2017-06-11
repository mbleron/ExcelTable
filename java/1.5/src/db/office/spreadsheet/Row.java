package db.office.spreadsheet;

import java.util.ArrayList;

public class Row extends ArrayList<Cell> {

	private static final long serialVersionUID = 6615809380592083862L;
	private int ref;
	
	public Row(int rowRef, int cellCount) {
		super(cellCount);
		this.ref = rowRef;
	}
	
	public int getRef() {
		return this.ref;
	}
	
	public Cell[] getCells() {
		Cell[] cells = new Cell[this.size()];
		return this.toArray(cells);
	}
	
}
