package com.excel.writer;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFileWriter {
	private ExcelCreator fileTowrite;
	private String fileName;
	
	public ExcelFileWriter(ExcelCreator fileToWrite, String fileName) {
		super();
		this.fileTowrite = fileToWrite;
		this.fileName = fileName;
	}
	
	public void writeFile() throws IOException{
		FileOutputStream fileOut = new FileOutputStream(fileName);
		this.fileTowrite.getWorkBook().write(fileOut);
		fileOut.close();
	}
	
	
}
