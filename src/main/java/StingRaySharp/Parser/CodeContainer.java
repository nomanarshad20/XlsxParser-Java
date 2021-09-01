package StingRaySharp.Parser;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CodeContainer {
	
	
	
	
	public static void createXlsxFile() throws IOException {

		String OutFileResult = "C:\\Users\\noman\\Desktop\\FIg\\" + new Date().getSeconds() + "_LienSummaryResult.xlsx";

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("TelosLegalCorp");

		// sheet.setColumnWidth(0, 6000);
		// sheet.setColumnWidth(1, 4000);

		List<String> headers = new ArrayList<>();
		headers.add("#");
		headers.add("Debtor Name");
		headers.add("Jurisdiction");
		headers.add("File No and Date");
		headers.add("Secured Party");
		headers.add("Colleteral");
		headers.add("Search Thru Date");
		headers.add("Status");

		CellStyle headerStyle = workbook.createCellStyle();
		// headerStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
		// headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font = workbook.createFont();
		font.setFontName("Arial");
		// font.setFontHeightInPoints((short) 16);
		// headerStyle.setWrapText(true);

		font.setBold(true);
		headerStyle.setFont(font);
		sheet.setDefaultColumnWidth(20);

		CellStyle titalStyle = workbook.createCellStyle();
		titalStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
		titalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		titalStyle.setFont(font);
		titalStyle.setAlignment(HorizontalAlignment.CENTER);
		titalStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		Row header = sheet.createRow(0);
		// sheet.autoSizeColumn(0);

		for (int i = 0; i < headers.size(); i++) {
			String headName = headers.get(i);
			Cell headerCell = header.createCell(i);
			headerCell.setCellValue(headName);

			headerCell.setCellStyle(headerStyle);
		}

		Map<String, Map<String, String>> rowDataMap = new HashMap<>();
		Map<String, String> rowInnerDataMap = new HashMap<>();
		rowInnerDataMap.put("#", "100");
		rowInnerDataMap.put("Debtor Name", "usa company");
		rowInnerDataMap.put("Jurisdiction", "super user");
		rowInnerDataMap.put("File No and Date", "28-3-90");
		rowInnerDataMap.put("Secured Party", "securing targer");
		rowInnerDataMap.put("Collateral", "people");
		rowInnerDataMap.put("Search Thru Date", "22-33-44");
		rowInnerDataMap.put("Status", "active");
		rowDataMap.put("Big Data", rowInnerDataMap);

		int rowInt = 1;

		for (Entry<String, Map<String, String>> rowDataSet : rowDataMap.entrySet()) {
			Row row = sheet.createRow(rowInt);
			sheet.autoSizeColumn(rowInt);

			sheet.addMergedRegion(new CellRangeAddress(rowInt, rowInt, 0, headers.size()));

			Cell headingCol = row.createCell(0);
			headingCol.setCellValue(rowDataSet.getKey());
			headingCol.setCellStyle(titalStyle);
			row.setHeightInPoints(40);

			rowInt++;

			int cellInt = 0;

			Row rowActualDataValue = sheet.createRow(rowInt);

			Map<String, String> innerDataColMap = rowDataSet.getValue();

			for (String fetchFromHeader : headers) {
				Cell dataCell = rowActualDataValue.createCell(cellInt);
				dataCell.setCellValue(innerDataColMap.get(fetchFromHeader));
				cellInt++;
			}
		}

		FileOutputStream outputStream = new FileOutputStream(OutFileResult);
		workbook.write(outputStream);
		workbook.close();
		System.out.println("***************8");

	}

	public static void createXlsxFile2() throws IOException {

		String OutFileResult = "C:\\Users\\noman\\Desktop\\FIg\\LienSummaryResult.xlsx";

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("TelosLegalCorp");

		sheet.setColumnWidth(0, 6000);
		sheet.setColumnWidth(1, 4000);

		Row header = sheet.createRow(0);

		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		List<String> headers = new ArrayList<>();
		headers.add("#");
		headers.add("Debtor Name");
		headers.add("Jurisdiction");
		headers.add("File No and Date");
		headers.add("Secured Party");
		headers.add("Collateral");
		headers.add("Search Thru Date");
		headers.add("Status");

		for (int i = 0; i < headers.size(); i++) {
			String headName = headers.get(i);
			Cell headerCell = header.createCell(i);
			headerCell.setCellValue(headName);
		}

		XSSFFont font = workbook.createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 16);
		font.setBold(true);
		headerStyle.setFont(font);

		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Name");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Age");
		headerCell.setCellStyle(headerStyle);

		CellStyle style = workbook.createCellStyle();
		style.setWrapText(true);

		Row row = sheet.createRow(2);
		Cell cell = row.createCell(0);
		cell.setCellValue("John Smith");
		cell.setCellStyle(style);

		cell = row.createCell(1);
		cell.setCellValue(20);
		cell.setCellStyle(style);

		// File currDir = new File(OutFileResult);
		// String path = currDir.getAbsolutePath();
		// String fileLocation = path.substring(0, path.length() - 1) +
		// "temp.xlsx";

		FileOutputStream outputStream = new FileOutputStream(OutFileResult);
		workbook.write(outputStream);
		workbook.close();
		System.out.println("***************8");

	}
/*
	
	
	private static void createXlsxFile(Map<String, List<Map<String, String>>> finalRowDataMap , List<String> orderList) throws IOException {

		String OutFileResult = "C:\\Users\\noman\\Desktop\\FIg\\" + new Date().getSeconds() + "_LienSummaryResul"
				+ ".xlsx";

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("TelosLegalCorp");

		List<String> headers = new ArrayList<>();
		headers.add("#");
		headers.add("Debtor Name");
		headers.add("Jurisdiction");
		headers.add("File No and Date");
		headers.add("Secured Party");
		headers.add("Collateral");
		headers.add("Search Thru Date");
		headers.add("Status");

		CellStyle headerStyle = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontName("Arial");
		headerStyle.setWrapText(true);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		font.setBold(true);
		headerStyle.setFont(font);

		CellStyle titalStyle = workbook.createCellStyle();
		titalStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		titalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		titalStyle.setFont(font);
		titalStyle.setWrapText(true);
		titalStyle.setAlignment(HorizontalAlignment.CENTER);
		titalStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		CellStyle rowStyle = workbook.createCellStyle();
		rowStyle.setAlignment(HorizontalAlignment.CENTER);
		rowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		rowStyle.setWrapText(true);

		sheet.setDefaultRowHeightInPoints(60);

		Row header = sheet.createRow(0);
		for (int i = 0; i < headers.size(); i++) {
			String headName = headers.get(i);
			Cell headerCell = header.createCell(i);
			headerCell.setCellValue(headName);
			headerCell.setCellStyle(headerStyle);
			// sheet.autoSizeColumn(i);
			// sheet.setColumnWidth(i,sheet.getColumnWidth(i)*17/10);
		}

		int rowInt = 0;
		
		
		
			
			
		
		
		
		for (String rowOrder : orderList) {
	//	for (Entry<String, List<Map<String, String>>> mainData : finalRowDataMap.entrySet()) {

			List<Map<String, String>> mainData = finalRowDataMap.get(rowOrder);
			
			rowInt++;
			Row row = sheet.createRow(rowInt);
			sheet.addMergedRegion(new CellRangeAddress(rowInt, rowInt, 0, headers.size()));
			Cell headingCol = row.createCell(0);
			headingCol.setCellValue(mainData.getKey());
			headingCol.setCellStyle(titalStyle);

			List<Map<String, String>> inerMapList = mainData.getValue();
			System.out.println(mainData.getKey() + "      " + inerMapList.toString());
			System.out.println(inerMapList.size());
			System.out.println(" ");

			for (Map<String, String> innerDataColMap : inerMapList) {

				rowInt++;
				Row rowActualDataValue = sheet.createRow(rowInt);
				int cellInt = 0;
				for (String fetchFromHeader : headers) {

					Cell dataCell = rowActualDataValue.createCell(cellInt);
					dataCell.setCellValue(innerDataColMap.get(fetchFromHeader));
					dataCell.setCellStyle(rowStyle);
					cellInt++;

					// wirth set
					sheet.autoSizeColumn(cellInt);
					sheet.setColumnWidth(cellInt, sheet.getColumnWidth(cellInt));
				}
			}
		}

		FileOutputStream outputStream = new FileOutputStream(OutFileResult);
		workbook.write(outputStream);
		workbook.close();
		System.out.println("********* end ******");

	}
	
	
	
	
	
*/
}
