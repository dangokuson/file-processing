package com.saltlux.fileprocessing;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.junit.Test;

public class XLS2CSVmraTest {

	@Test
	public void testXLS2CSVmra01() throws FileNotFoundException, IOException {
		long start = System.currentTimeMillis();
		final String fileName = getClass().getClassLoader().getResource("EastNorthCentralFTWorkers.xls").getPath();
		
		XLS2CSVmra xls2csv = new XLS2CSVmra(fileName, "Data", -1);
		xls2csv.process();
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");
	}

}
