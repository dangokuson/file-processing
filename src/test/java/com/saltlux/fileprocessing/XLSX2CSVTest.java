package com.saltlux.fileprocessing;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.junit.Test;
import org.xml.sax.SAXException;

public class XLSX2CSVTest {

	@Test
	public void test() throws IOException, OpenXML4JException, SAXException {
		long start = System.currentTimeMillis();
		String filePath = getClass().getClassLoader().getResource("test08.XLSX").getPath();
		File xlsxFile = new File(filePath);
		
		// The package open is instantaneous, as it should be.
        OPCPackage p = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ);
		XLSX2CSV xlsx2csv = new XLSX2CSV(p, System.out, -1);
		xlsx2csv.process();
		p.close();
		System.out.println("Read in " + (System.currentTimeMillis() - start) + "ms");
	}

}
