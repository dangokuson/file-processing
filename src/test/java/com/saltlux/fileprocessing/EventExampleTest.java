package com.saltlux.fileprocessing;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.Test;

public class EventExampleTest {

	@Test
	public void testHSSFEvent() throws IOException {
		long start = System.currentTimeMillis();
		final String fileName = getClass().getClassLoader().getResource("EastNorthCentralFTWorkers.xls").getPath();
		// create a new file input stream with the input file specified
        // at the command line
        FileInputStream fin = new FileInputStream(new File(fileName));
        // create a new org.apache.poi.poifs.filesystem.Filesystem
        POIFSFileSystem poifs = new POIFSFileSystem(fin);
        // get the Workbook (excel part) stream in a InputStream
        InputStream din = poifs.createDocumentInputStream("Workbook");
        // construct out HSSFRequest object
        HSSFRequest req = new HSSFRequest();
        // lazy listen for ALL records with the listener shown above
        req.addListenerForAllRecords(new EventExample());
        // create our event factory
        HSSFEventFactory factory = new HSSFEventFactory();
        // process our events based on the document input stream
        factory.processEvents(req, din);
        // once all the events are processed close our file input stream
        fin.close();
        // and our document input stream (don't want to leak these!)
        din.close();
        System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");
	}

}
