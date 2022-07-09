package com.Apache_POI.com.Apache_POI;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteIntoExcel {

	public static void main(String[] args) throws IOException {

		// Blank workbook
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Sheet Name");

		// This data needs to be written (Object[])
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "ID", "FIRSTNAME", "LASTNAME" });
		data.put("2", new Object[] { 1, "Anuj", "Malik" });
		data.put("3", new Object[] { 2, "Shwetank", "Vishnu" });
		data.put("4", new Object[] { 3, "Vishal", "Saini" });
		data.put("5", new Object[] { 4, "Prashant", "Gangwar" });
		data.put("6", new Object[] { 4, "Ranvijay", "Singh" });

		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}

		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File("Filename.xlsx"));
		workbook.write(out);
		out.close();
		System.out.println("written successfully in File.");

		// Run this class and Refresh Project

	}

}
