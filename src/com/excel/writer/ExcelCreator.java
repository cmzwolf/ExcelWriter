package com.excel.writer;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;

public class ExcelCreator {

	private static final String NumberLabel = "number";
	private static final String DateLabel = "date";

	private Boolean roundNumerics;
	private String sheetName;
	List<String> columnsNames;
	private Map<String, String> columnsTypes;
	private Map<String, List<String>> columnsDataContent;
	private HSSFWorkbook workBook;

	private CellStyle headerCellStyle;
	private CellStyle genericCellStyle;
	private CellStyle numericCellStyle;
	private CellStyle dateCellStyle;

	public ExcelCreator(String sheetName, List<String> columnsNames,
			Map<String, String> columnsTypes,
			Map<String, List<String>> columnsDataContent, Boolean roundNumerics) {
		super();
		this.sheetName = sheetName;
		this.columnsNames = columnsNames;
		this.columnsTypes = columnsTypes;
		this.columnsDataContent = columnsDataContent;
		this.roundNumerics = roundNumerics;
		this.createExcelFile();
	}

	private void initializeStyles() {
		this.workBook = new HSSFWorkbook();
		CreationHelper createHelper = workBook.getCreationHelper();

		genericCellStyle = workBook.createCellStyle();
		genericCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// defining styles for number format manipulation
		numericCellStyle = workBook.createCellStyle();
		if (this.roundNumerics) {
			DataFormat numberFormat = workBook.createDataFormat();
			numericCellStyle
					.setDataFormat(numberFormat.getFormat("#,##0.0000"));
		}
		numericCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// defining styles for date format manipulation
		dateCellStyle = workBook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"dd/mm/yyyy"));
		dateCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		//define the style for the Header
		headerCellStyle = workBook.createCellStyle();
		headerCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFFont font = workBook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Arial");
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setItalic(false);
		headerCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerCellStyle.setFont(font);
	}

	public HSSFWorkbook getWorkBook() {
		return workBook;
	}

	private void createExcelFile() {
		this.initializeStyles();

		HSSFSheet sheet = workBook.createSheet(sheetName);
		Integer columnsNumber = columnsNames.size();

		Integer maxDataLength = computeRowMaxNumber();

		for (int i = 0; i < maxDataLength; i++) {
			HSSFRow currentRow = sheet.createRow(i + 1);
			for (int j = 0; j < columnsNumber; j++) {
				try {
					String columnName = columnsNames.get(j);
					String columnType = columnsTypes.get(columnName);
					String cellContent = columnsDataContent.get(columnName)
							.get(i);
					this.cellFactory(currentRow, j, columnType, cellContent,
							genericCellStyle);
				} catch (Exception e) {
					// do nothing
					// one can not just write the cell into the excel
				}

			}
		}

		// writing the header line
		HSSFRow headerRow = sheet.createRow(0);
		for (int i = 0; i < columnsNumber; i++) {
			String columnName = columnsNames.get(i);
			HSSFCell cell = headerRow.createCell(i);
			cell.setCellValue(columnName);
			cell.setCellStyle(headerCellStyle);
			sheet.autoSizeColumn(i);
		}
	}

	private HSSFCell cellFactory(HSSFRow row, Integer cellPosition,
			String cellType, String cellValue, CellStyle cellStyle) {
		HSSFCell cell = null;

		if (cellType.equalsIgnoreCase(NumberLabel)) {
			cell = row.createCell(cellPosition, Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Double.parseDouble(cellValue));
			cell.setCellStyle(numericCellStyle);
		} else {
			if (cellType.equalsIgnoreCase(DateLabel)) {
				try {
					// defining styles for date format manipulation
					String dateFormatString = "yyyyMMdd";
					SimpleDateFormat dateFormat = new SimpleDateFormat(
							dateFormatString);
					Date date = null;
					date = dateFormat.parse(cellValue);
					cell = row.createCell(cellPosition);
					cell.setCellStyle(dateCellStyle);
					cell.setCellValue(date);

				} catch (ParseException e) {
					e.printStackTrace();
				}
			} else {
				cell = row.createCell(cellPosition, Cell.CELL_TYPE_STRING);
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}
		}
		return cell;
	}

	private Integer computeRowMaxNumber() {
		Integer toReturn = -1;
		Integer sizeOfData;
		for (Entry<String, List<String>> entry : columnsDataContent.entrySet()) {
			sizeOfData = entry.getValue().size();
			if (sizeOfData >= toReturn) {
				toReturn = sizeOfData;
			}
		}
		return toReturn;
	}

}
