package com.excel.writer.example;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.excel.writer.ExcelCreator;
import com.excel.writer.ExcelFileWriter;

public class ExampleofUse {
	public static void main(String[] args) throws IOException {

		// preparing data structure

		String sheetName = "test";
		

		List<String> columnsNames = new ArrayList<String>();
		columnsNames.add("firstColumn");
		columnsNames.add("secondColumn");
		columnsNames.add("thirdColumn");
		columnsNames.add("fourthColumn");

		Map<String, String> columnsTypes = new HashMap<String, String>();
		columnsTypes.put(columnsNames.get(0), "number");
		columnsTypes.put(columnsNames.get(1), "number");
		columnsTypes.put(columnsNames.get(2), "String");
		columnsTypes.put(columnsNames.get(3), "date");
		
		
		Map<String, List<String>> columnsDataContent = new HashMap<String, List<String>>();
		for (int i = 0; i < columnsNames.size(); i++) {
			List<String> tempList = new ArrayList<String>();
			for (int j = 0; j < 100; j++) {
				tempList.add(new Double(100*Math.random()).toString());
			}
			columnsDataContent.put(columnsNames.get(i), tempList);
		}
		
		String datePattern="199901";
		
		List<String> dates = new ArrayList<String>();
		for(int i=10; i<20;i++){
			dates.add(datePattern+""+i);
		}
		
		columnsDataContent.put(columnsNames.get(3), dates);
		

		ExcelCreator ec = new ExcelCreator(sheetName, columnsNames,
				columnsTypes, columnsDataContent, false);
		
		ExcelFileWriter fl = new ExcelFileWriter(ec, "myExcelFile.xls");
		fl.writeFile();

	}


}
