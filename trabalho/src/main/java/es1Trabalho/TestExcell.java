package es1Trabalho;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestExcell {

	public void writeExcell() {

		try {
			Workbook workbook = new XSSFWorkbook();

			Sheet sh = workbook.createSheet("Invoices");

			String[] collumHeadings = { "Item id", "Item Name", "Qty", "Item Price", "Sold Date" };

			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 12);
			headerFont.setColor(IndexedColors.BLACK.index);

			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			
			Row headerRow = sh.createRow(0);
			
			for (int i=0; i<collumHeadings.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(collumHeadings[i]);
				cell.setCellStyle(headerStyle);
			}

			//ArrayList<E>
			
		} catch (Exception e) {

		}

	}
}
