package db.office.spreadsheet;

import java.io.IOException;
import java.util.List;

public interface CellReader {

	public int getColumnCount();
	public void addSheet(Sheet sheet);
	public List<String> getSheetList();
	public List<Row> readRows(int nrows) throws CellReaderException;
	public void close() throws CellReaderException, IOException;
	
}
