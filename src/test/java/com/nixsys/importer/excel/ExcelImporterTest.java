package com.nixsys.importer.excel;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

public class ExcelImporterTest {
	
	@Test
	public void shouldImportExcelSuccessfullyFromStaticFile() throws InvalidFormatException, IOException {
		ExcelImporter excelImporter = new ExcelImporter(new File("d:\\test.xlsx"));
		Map<Integer, Object[]> result = excelImporter.extractData();
		printConsoleResult(result);
		//TODO do the asserts when get the correct excel file
	}
	
	private void printConsoleResult(Map<Integer, Object[]> result) {
		for (Integer key : result.keySet()) {
			System.out.println("--------------------------------------------------------------------------------------------------------------------------");
			System.out.println("key : " + key);
			System.out.println("--------------------------------------------------------------------------------------------------------------------------");
			Object[] data = result.get(key);
			for (int i = 0; i < data.length; i++) {
				if (i <= data.length) {
					System.out.print(data[i] + " | ");
				} else {
					System.out.println(data[i] + " | ");
				}
			}
			System.out.println("\n");
		}
	}

}
