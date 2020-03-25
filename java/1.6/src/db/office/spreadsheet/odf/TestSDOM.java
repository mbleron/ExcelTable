package db.office.spreadsheet.odf;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import javax.xml.namespace.QName;
import oracle.xml.parser.v2.SAXParser;
import oracle.xml.scalable.BinaryStream;
import oracle.xml.scalable.InfosetReader;

public class TestSDOM {
	
	static final String ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
	static final QName QNAME_TABLE = new QName(ODF_TABLE_NS, "table");
	static final QName QNAME_TNAME = new QName(ODF_TABLE_NS, "name");
	static final QName QNAME_TROW = new QName(ODF_TABLE_NS, "table-row");
	
	public static void main(String[] args) throws Exception {
		
		long startTime = System.nanoTime();
		
		FileInputStream is = new FileInputStream("/dev/oracle/tmp/xl_data/content.xml");
		ByteArrayOutputStream out = new ByteArrayOutputStream(4096);
		
        SAXParser parser = new SAXParser();
        parser.setValidationMode(SAXParser.NONVALIDATING);
        parser.setPreserveWhitespace(false);
        
        SAXDocumentSerializer serializer = new SAXDocumentSerializer();
        serializer.setOutputStream(out);
        
        parser.setContentHandler(serializer);
        //parser.setProperty("http://xml.org/sax/properties/lexical-handler", serializer);
        parser.parse(is);
        is.close();

		System.out.println((System.nanoTime() - startTime)/1e9);
		
		BinaryStream stream = BinaryStream.newInstance(BinaryStream.SUN_FI);
		stream.setByteArray(out.toByteArray());
		
		//System.out.println(out.size());
		
		int rowCount = 0;
		
		InfosetReader reader = stream.getInfosetReader();
		
		while (reader.hasNext()) {
			reader.next();
			int event = reader.getEventType();
			if (event == InfosetReader.START_ELEMENT) {	
				if (reader.getQName().equals(QNAME_TABLE)) {
					rowCount = 0;
					System.out.println(reader.getAttributes().getValue(QNAME_TNAME.getNamespaceURI(), QNAME_TNAME.getLocalPart()));
				} else if (reader.getQName().equals(QNAME_TROW)) {
					rowCount++;
					System.out.println(rowCount);
				}
			}
		}
		
		reader.close();
		stream.close();
		
		long estimatedTime = System.nanoTime() - startTime;
		System.out.println(estimatedTime/1e9);

	}

}
