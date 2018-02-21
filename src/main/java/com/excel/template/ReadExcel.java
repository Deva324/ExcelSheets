package com.excel.template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	private static final String fileName = "C:\\Users\\dgadiam\\Documents\\";

	public static void main(String[] args) {

		String name = "Sprint-StoryTracker-Template-Input.xlsx";

		Map<String, String> stories = ExcelUtil.readFile(name);
		generateExcel(stories);

	}

	private static void generateExcel(Map<String, String> stories) {
		try {
			FileUtils.copyFile(new File(fileName + "Sprint-StoryTracker-Template-dup.xlsx"),
					new File(fileName + "Sprint-StoryTracker-Template2.xlsx"));
			FileInputStream fis = new FileInputStream(fileName + "Sprint-StoryTracker-Template2.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet s = workbook.getSheet("Overview");
			CellCopyPolicy cp = new CellCopyPolicy();

			for (Map.Entry<String, String> entry : stories.entrySet()) {
				XSSFSheet storySheet = workbook.cloneSheet(1, entry.getKey().replace("-", ""));
				storySheet.getRow(0).getCell(1).setCellValue(entry.getValue());
				int n = s.getLastRowNum();
				XSSFRow srcRow = s.getRow(1);
				XSSFRow newRow = s.createRow(n + 1);
				s.copyRows(1, 1, n + 1, cp);
				newRow = s.getRow(n + 1);
				newRow.getCell(0).setCellValue(entry.getKey());
				newRow.getCell(1).setCellValue(entry.getValue());
				ExcelUtil.copyRow(srcRow, newRow, storySheet.getSheetName());
			}

			/*
			 * XSSFDrawing d = s.getDrawingPatriarch(); List<XSSFChart> l = d.getCharts();
			 * for (int i = 0; i < l.size(); i++) { XSSFChart c = l.get(i); //
			 * System.out.println(c.getCTChart()); List<CTBarChart> c1 =
			 * c.getCTChart().getPlotArea().getBarChartList(); for (CTBarChart c2 : c1) {
			 * List<CTBarSer> s2 = c2.getSerList(); for (CTBarSer s3 : s2) {
			 * System.out.println(s3.getTx().getStrRef().getF());
			 * 
			 * System.out.println(s3.getVal().getNumRef().getF());
			 * 
			 * } }
			 * 
			 * }
			 */
			/*
			 * XSSFDrawing d1 = workbook.cloneSheet(1, "Test").createDrawingPatriarch();
			 * List<XSSFChart> l1 = d1.getCharts(); for (int i = 0; i < l1.size(); i++) {
			 * XSSFChart c = l1.get(i); List<CTBarChart> c1 =
			 * c.getCTChart().getPlotArea().getBarChartList(); for (CTBarChart c2 : c1) {
			 * List<CTBarSer> s2 = c2.getSerList(); for (CTBarSer s3 : s2) { String
			 * s4=s3.getTx().getStrRef().getF(); s4=s4.replace("StoryID", "Test");
			 * s3.getTx().getStrRef().setF(s4); String s5=s3.getVal().getNumRef().getF();
			 * s5=s5.replace("StoryID", "Test"); s3.getVal().getNumRef().setF(s5);
			 * //System.out.println(s4); //System.out.println(s5);
			 * //System.out.println(s3.getTx().getStrRef().getF());
			 * //System.out.println(s3.getVal().getNumRef().getF());
			 * 
			 * } } System.out.println(l1.get(i).getCTChart()); }
			 */
			FileOutputStream fileOut = new FileOutputStream(new File(fileName + "Sprint-StoryTracker-Template2.xlsx"));
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

}
