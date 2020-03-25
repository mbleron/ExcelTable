package db.office.spreadsheet;

import java.sql.Blob;

public class Sheet {

	private Blob content;
	private int index = 0;
	
	public Sheet(int index) {
		this.index = index;
	}
	
	public Sheet(int index, Blob content) {
		this.index = index;
		this.content = content;
	}
	
	public int getIndex() {
		return this.index;
	}
	
	public Blob getContent() {
		return this.content;
	}
	
}
