package com.bigBasket.uploadTestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.instrument.IllegalClassFormatException;
import java.text.SimpleDateFormat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelLibrary {
	String filePath;

	public ExcelLibrary(String filepath) {
		this.filePath = filepath;
	}

	public Sheet getSheet() throws InvalidFormatException, IOException {
		FileInputStream fis = new FileInputStream(filePath);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheetAt(0);
		return sh;
	}

	@SuppressWarnings("deprecation")
	public String getExcelData(int rowNo, int colNo) throws InvalidFormatException, IOException {
		String strCellValue = null;
		Sheet sh = getSheet();
		Row row = sh.getRow(rowNo);
		if (row == null) return "";
		Cell cell = row.getCell(colNo);
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				strCellValue = cell.toString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
					strCellValue = dateFormat.format(cell.getDateCellValue().toString());
				} else {
					Double value = cell.getNumericCellValue();
					Long longValue = value.longValue();
					strCellValue = new String(longValue.toString());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				strCellValue = new String(new Boolean(cell.getBooleanCellValue()).toString());
				break;
			case Cell.CELL_TYPE_BLANK:
				strCellValue = "";
				break;
			}
		}
		return strCellValue;
	}

	public int getRowCount() throws IllegalClassFormatException, IOException, InvalidFormatException {
		Sheet sh = getSheet();
		//int rowCount = sh.getLastRowNum();
		int rowCount = sh.getPhysicalNumberOfRows();
		return rowCount;
	}

	public int getColumnCount() throws IllegalClassFormatException, IOException, InvalidFormatException {
		Sheet sh = getSheet();
		int columnCount = sh.getRow(0).getLastCellNum();
		return columnCount;
	}

	@SuppressWarnings("deprecation")
	public void setData(int rowNo, int colNo, String data) throws InvalidFormatException, IOException {
		FileInputStream fis = new FileInputStream(filePath);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = getSheet();
		Row row = sh.getRow(rowNo);
		Cell cell = row.createCell(colNo);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		FileOutputStream fos = new FileOutputStream(filePath);
		cell.setCellValue(data);
		wb.write(fos);

	}
}