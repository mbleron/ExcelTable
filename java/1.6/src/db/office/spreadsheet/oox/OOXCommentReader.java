package db.office.spreadsheet.oox;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import db.office.spreadsheet.CellReaderException;

public class OOXCommentReader {

	private static final String TAG_COMMENT = "comment";
	private static final String TAG_TEXT = "text";
	private static final String TAG_T = "t";
	private static final String TAG_REF = "ref";
	
	private static XMLInputFactory factory = XMLInputFactory.newInstance();
	private XMLStreamReader reader;
	private String sheetRef;
	
	public OOXCommentReader(int sheetIndex, InputStream is) throws XMLStreamException {
		this.sheetRef = String.valueOf(sheetIndex);
		reader = factory.createXMLStreamReader(is);
	}
	
	private Map<String,String> readComments() throws XMLStreamException {	
		Map<String,String> comments = new HashMap<String,String>();	
		while (reader.hasNext()) {
			int event = reader.next();
			if (event == XMLStreamReader.START_ELEMENT && reader.getLocalName().equals(TAG_COMMENT)) {					
				String cellRef = reader.getAttributeValue(null, TAG_REF);
				String text = readText();
				comments.put(sheetRef + cellRef, text);
			}
		}	
		reader.close();	
		return comments;	
	}
	
	private String readText() throws XMLStreamException {
		
		boolean in_text = false;
		StringBuilder sb = new StringBuilder();
		
		while (reader.hasNext()) {
			int event = reader.next();
			switch (event) {
			case XMLStreamReader.START_ELEMENT:
				if (reader.getLocalName().equals(TAG_TEXT)) {
					in_text = true;
				} else if (in_text && reader.getLocalName().equals(TAG_T)) {
					sb.append(reader.getElementText());
				}
				break;
			case XMLStreamReader.END_ELEMENT:
				if (reader.getLocalName().equals(TAG_TEXT)) {
					return sb.toString();
				}
			}
		}
		
		throw new IllegalStateException();
		
	}
	
	public static Map<String,String> getComments(int sheetIndex, InputStream is) throws CellReaderException, IOException {
		try {
			OOXCommentReader reader = new OOXCommentReader(sheetIndex, is);
			return reader.readComments();
		} catch (Exception e) {
			throw new CellReaderException("Error in getComments() method", e);
		} finally {
			is.close();
		}
	}
	
	/*
	public static void main(String[] args) throws Exception {
		FileInputStream is = new FileInputStream(args[0]);
		Map<String,String> comments = OOXCommentReader.getComments(is);
		for (Map.Entry<String, String> entry : comments.entrySet()) {
			System.out.println(entry.toString());
		}
		is.close();
	}
	*/
	
}
