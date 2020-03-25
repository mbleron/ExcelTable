package db.office.spreadsheet.odf;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import javax.xml.stream.XMLStreamException;
import db.office.spreadsheet.Cell;
import db.office.spreadsheet.CellReaderException;
import db.office.spreadsheet.Row;
import db.office.spreadsheet.Sheet;

public class TestImpl {

	public static void main(String[] args) throws CellReaderException, IOException, XMLStreamException  {
		
		FileInputStream is = new FileInputStream("/dev/oracle/tmp/xl_data/test01/content.xml");
		ODFCellReaderImpl cellReader = new ODFCellReaderImpl("A,B,C,D,E", 1, 100);
		cellReader.setContent(is);
		
		//cellReader.readAll();
		
		
		cellReader.addSheet(new Sheet(1));
		//cellReader.addSheet(2);
		//cellReader.nextSheet();
		
		
		List<Row> rows = cellReader.readRows(100);
		
		for (Row r:rows) {
			for (Cell<?> c:r) {
				System.out.println(c);
			}
		}
		
		cellReader.close();

	}

}
