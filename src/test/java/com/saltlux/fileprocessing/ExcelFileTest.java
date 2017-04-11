package com.saltlux.fileprocessing;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class ExcelFileTest {
	 
	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
	}
 
	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
	}
 
	/**
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws SAXException
	 */
	@Test
	public void testReadBigExcelFile() throws IOException, OpenXML4JException,
			SAXException {
		long start = System.currentTimeMillis();
		final String fileName = getClass().getClassLoader().getResource("Sample-SuperstoreSubset.xlsx").getPath();
		/*
		 * Opens a package (archive / xlsx file) with read / write permissions.
		 * It is also possible to access it read only, which should be the first
		 * choice for read operations in case the file is already accessed by
		 * another user. To open read only provide an InputStream instead of a
		 * file path.
		 */
		OPCPackage pkg = OPCPackage.open(fileName);
		XSSFReader r = new XSSFReader(pkg);
 
		/*
		 * Read the sharedStrings.xml from an xlsx file into an object
		 * representation.
		 */
		SharedStringsTable sst = r.getSharedStringsTable();
 
		/*
		 * Hand a read SharedStringsTable for further reference to the SAXParser
		 * and the underlying ContentHandler.
		 */
		XMLReader parser = fetchSheetParser(sst);
 
		/*
		 * To look up the Sheet Name / Sheet Order / rID, you need to process
		 * the core Workbook stream. Normally it's of the form rId# or rSheet#
		 * 
		 * Great! How do I know, if it is rSheet or rId? Thanks Microsoft.
		 * Anyhow, let's carry on with the noise.
		 * 
		 * I reference the third sheet from the left since the index starts at
		 * one.
		 */
		InputStream sheet2 = r.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet2);
 
		/*
		 * Run through a Sheet using a window of several XML tags instead of
		 * attempting to read the whole file into RAM at once. Leaves the
		 * handling of file content to the ContentHandler, which is in this case
		 * the nested class SheetHandler.
		 */
		parser.parse(sheetSource);
 
		/*
		 * Close the underlying InputStream for a Sheet XML.
		 */
		sheet2.close();
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");
	}
 
	public XMLReader fetchSheetParser(SharedStringsTable sst)
			throws SAXException {
		/*
		 * XMLReader parser = XMLReaderFactory
		 * .createXMLReader("org.apache.xerces.parsers.SAXParser");
		 */
		XMLReader parser = XMLReaderFactory.createXMLReader();
 
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}
 
	private static class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
 
		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}
 
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {
			// c => cell
			if (name.equals("c")) {
				// Print the cell reference
				System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if (cellType != null && cellType.equals("s")) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			// Clear contents cache
			lastContents = "";
		}
 
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx))
						.toString();
				nextIsString = false;
			}
 
			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v")) {
				System.out.println(lastContents);
			}
		}
 
		public void characters(char[] ch, int start, int length)
				throws SAXException {
			lastContents += new String(ch, start, length);
		}
	}
 
}
