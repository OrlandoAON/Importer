package com.nixsys.importer.excel;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImporter {
	
	private static final int FIRST_POSITION = 0;

	private XSSFWorkbook workbook;
	
	private Integer columns;
	
	private Integer rows;
	
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public Integer getColumns() {
		return columns;
	}

	public void setColumns(Integer columns) {
		this.columns = columns;
	}

	public Integer getRows() {
		return rows;
	}

	public void setRows(Integer rows) {
		this.rows = rows;
	}
	
	public ExcelImporter(File file) throws InvalidFormatException, IOException {
		this.workbook = new XSSFWorkbook(file);
		init();
	}

	private void init() {
		
		XSSFSheet sheet = getWorkbook().getSheetAt(0);
		int maxColumn = 0;
		int maxRow = 0;
		Iterator<Row> rows = sheet.iterator();
		Iterator<Cell> header = rows.next().cellIterator();
		if (header != null) {
			maxRow++;
		}
		while (header.hasNext()) {
			Cell cell = header.next();
			if (!StringUtils.isBlank(cell.getStringCellValue())) {
				maxColumn++;
			}
		}
		while (rows.hasNext()) {
			Row row = rows.next();
			Cell firstRow = row.getCell(FIRST_POSITION);
			if (firstRow == null) {
				break;
			}
			if (!StringUtils.isBlank(firstRow.getStringCellValue())) {
				maxRow++;
			}
		}
		setColumns(maxColumn);
		setRows(maxRow);
	}
	
	public Map<Integer, Object[]> extractData() {
		
		Map<Integer, Object[]> result = new HashMap<Integer, Object[]>();
		XSSFSheet sheet = getWorkbook().getSheetAt(0);
		Iterator<Row> rows = sheet.iterator();
		rows.next();//skip header
		while (rows.hasNext()) {
			Row row = rows.next();
			if (row.getRowNum() > getRows()) {
				break;
			}
			Iterator<Cell> cells = row.cellIterator();
			Object[] data = new Object[getColumns()];
			result.put(row.getRowNum(), data);
			while (cells.hasNext()) {
				Cell cell = cells.next();
				if (cell.getColumnIndex() <= getColumns()) {
					data[cell.getColumnIndex()] = cell.getStringCellValue(); 
				}
			}
		}
		return result;
	}

}
