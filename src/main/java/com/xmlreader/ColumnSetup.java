package com.xmlreader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColumnSetup extends Constants{

	/**
	 * Initializes the POI workbook and writes the header row
	 */
	public Workbook initXls() {
		Workbook workbook = new XSSFWorkbook();

		CellStyle style = workbook.createCellStyle();
		Font boldFont = workbook.createFont();
		boldFont.setBold(true);
		style.setFont(boldFont);
		style.setAlignment(CellStyle.ALIGN_CENTER);

		Sheet sheet = workbook.createSheet();
		int rowNum = 0;
		Row row = sheet.createRow(rowNum);
		
		Cell cell = row.createCell(0);
		/*
		 * cell.setCellValue(REGIMEN_FISCAL + EMISOR);
		 */
		cell.setCellStyle(style);

		cell = row.createCell(1);
		cell.setCellValue(RFC_EMISOR + EMISOR);
		cell.setCellStyle(style);

		cell = row.createCell(2);
		cell.setCellValue(NOMBRE_EMISOR + EMISOR);
		cell.setCellStyle(style);

		cell = row.createCell(3);
		cell.setCellValue(RFC_RECEPTOR + RECEPTOR);
		cell.setCellStyle(style);

		cell = row.createCell(4);
		cell.setCellValue(NOMBRE_RECEPTOR + RECEPTOR);
		cell.setCellStyle(style);

		cell = row.createCell(5);
		cell.setCellValue(USO_CFDI + RECEPTOR);
		cell.setCellStyle(style);

		cell = row.createCell(6);
		cell.setCellValue("Application Date");
		cell.setCellStyle(style);
		
		return workbook;

	}
	
}
