package db.office.spreadsheet;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

public class CellReader {

	private static final String TAG_ROW = "row";
	private static final String TAG_C = "c";
	private static final String TAG_V = "v";
	private static final String TAG_T = "t";
	//private static final String TAG_IS = "is";
	private static final String TAG_R = "r";
	private static final String CT_SHAREDSTRING = "s";
	//private static final String CELLTYPE_STR = "str";
	private static final String CT_INLINESTR = "inlineStr";
	
	private boolean done = false;
	private int firstRow = 0;
	private int lastRow;
	private InputStream source;
	private XMLStreamReader reader;
	private String[] strings;
	private Set<String> columns;
	
	public CellReader(InputStream worksheet, InputStream sharedStrings, String columnList, int firstRow, int lastRow)
			throws CellReaderException, IOException {
		
		this.source = worksheet;
		XMLInputFactory factory = XMLInputFactory.newInstance();
		//factory.setProperty(XMLInputFactory.IS_COALESCING, Boolean.TRUE);
		try {
			this.reader = factory.createXMLStreamReader(worksheet);
			if (sharedStrings != null) {
				this.strings = SharedStringsHandler.getStrings(sharedStrings);
			}
		} catch (Exception e) {
			throw new CellReaderException("Error during CellReader initialization", e);
		}
			
		this.columns = new HashSet<String>(Arrays.asList(columnList.split(",")));
		this.lastRow = lastRow;
		this.readRows(1, true, firstRow);	
		
	}
	
	public void close () throws CellReaderException, IOException {
		try {
			this.reader.close();
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error while closing CellReader", e);
		} finally {
			this.source.close();
		}
	}
	
	public int getColumnCount () {
		return this.columns.size();
	}	
	
	public List<Row> readRows(int nrows, boolean skipMode, int startWith) throws CellReaderException {

		List<Row> rows = null;

		try {

			if (!skipMode) {
				rows = new ArrayList<Row>(nrows);
				if (this.firstRow != 0) {
					rows.add(readRow(this.firstRow));
					this.firstRow = 0;
					nrows--;
				}
			}

			while (!this.done && this.reader.hasNext() && nrows > 0) {
				int event = this.reader.next();
				if (event == XMLStreamReader.START_ELEMENT 
						&& this.reader.getLocalName().equals(TAG_ROW)) {
					int rowRef = Integer.parseInt(this.reader.getAttributeValue(null, TAG_R));

					if (!skipMode) {
						rows.add(readRow(rowRef));
						nrows--;
						if (this.lastRow != -1 && rowRef >= this.lastRow) {
							done = true;
							break;
						}

					} else if (rowRef >= startWith) {
						this.firstRow = rowRef;
						break;
					}
				}
			}
			
			return rows;
			
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error in readRows method", e);
		}

	}	
	
	private Row readRow(int rowRef) throws XMLStreamException {
		
	    Row row = new Row(rowRef, this.columns.size());
	    Cell cell;
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						if ((cell = readCell()) != null) {
							row.add(cell);
						}
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_ROW)) {
						return row;
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}	
	
	private Cell readCell() throws XMLStreamException {

		CellRef cellRef = new CellRef(this.reader.getAttributeValue(null, TAG_R));
		Cell cell = null;
		
		if (columns.contains(cellRef.column)) {
		
			String cellType = this.reader.getAttributeValue(null, TAG_T);
			String cellValue = null;

			if (CT_INLINESTR.equals(cellType)) {
				// when t="inlineStr", <is> is the only child element allowed
				cellValue = readInlineStr();
			} else if (CT_SHAREDSTRING.equals(cellType)) {
				int idx = Integer.parseInt(readCellValue());
				cellValue = this.strings[idx];
			} else {
				cellValue = readCellValue();
			}

			cell = new Cell(cellRef, cellValue, cellType);

		} else {
			cell = new Cell(cellRef, "", "");
		}

		return cell;

	}
	
	private String readInlineStr() throws XMLStreamException {

		StringBuilder sb = new StringBuilder();
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_T)) {
						sb.append(this.reader.getElementText());
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						return sb.toString();
					}
			}	
		}
		
		throw new IllegalStateException();
		
	}
	
	private String readCellValue() throws XMLStreamException {
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_V)) {
						return this.reader.getElementText();
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						return "";
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}
	
}
