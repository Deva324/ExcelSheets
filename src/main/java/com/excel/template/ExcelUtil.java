package com.excel.template;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	private static final String fileName = "C:\\Users\\dgadiam\\Documents\\";

	public static Map<String, String> readFile(String name) {
		Map<String, String> stories = new HashMap<String, String>();
		try (FileInputStream fis = new FileInputStream(fileName + "Sprint-StoryTracker-Template-Input.xlsx")) {
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet s = workbook.getSheetAt(0);
			int noOfRows = s.getLastRowNum();
			for (int i = 1; i <= noOfRows; i++) {
				String storyNumber = s.getRow(i).getCell(0).getStringCellValue();
				String storyDescription = s.getRow(i).getCell(1).getStringCellValue();
				stories.put(storyNumber, storyDescription);
			}

			workbook.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return stories;
	}

	public static void copyRow(XSSFRow srcRow, XSSFRow newRow, String sheetName) {
		int j = srcRow.getFirstCellNum();
		if (j < 0) {
			j = 0;
		}
		for (; j <= srcRow.getLastCellNum(); j++) {
			XSSFCell oldCell = srcRow.getCell(j);
			XSSFCell newCell = newRow.getCell(j);
			if (oldCell != null) {
				if (oldCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
					String s4 = oldCell.getCellFormula().replace("StoryID", sheetName);
					newCell.setCellFormula(s4);
				}

			}

		}

	}

}
