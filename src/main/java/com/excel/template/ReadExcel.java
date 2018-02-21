package com.excel.template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;

public class ReadExcel {
	private static final String fileName = "C:\\Users\\dgadiam\\Documents\\";

	public static void main(String[] args) {

		String name = "Sprint-StoryTracker-Template-Input.xlsx";

		Map<String, String> stories = readFile(name);
		generateExcel(stories);

	}

	private static void generateExcel(Map<String, String> stories) {
		try {
			FileUtils.copyFile(new File(fileName + "Sprint-StoryTracker-Template.xlsx"),
					new File(fileName + "Sprint-StoryTracker-Template1.xlsx"));
			FileInputStream fis = new FileInputStream(fileName + "Sprint-StoryTracker-Template1.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet s = workbook.getSheet("StoryID");
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
			XSSFDrawing d1 = workbook.cloneSheet(1, "Test").createDrawingPatriarch();
			List<XSSFChart> l1 = d1.getCharts();
			for (int i = 0; i < l1.size(); i++) {
				XSSFChart c = l1.get(i);
				List<CTBarChart> c1 = c.getCTChart().getPlotArea().getBarChartList();
				for (CTBarChart c2 : c1) {
					List<CTBarSer> s2 = c2.getSerList();
					for (CTBarSer s3 : s2) {
						String s4=s3.getTx().getStrRef().getF();
						s4=s4.replace("StoryID", "Test");
						s3.getTx().getStrRef().setF(s4);
						String s5=s3.getVal().getNumRef().getF();
						s5=s5.replace("StoryID", "Test");
						s3.getVal().getNumRef().setF(s5);
						//System.out.println(s4);
						//System.out.println(s5);
						//System.out.println(s3.getTx().getStrRef().getF());
						//System.out.println(s3.getVal().getNumRef().getF());

					}
				}
				System.out.println(l1.get(i).getCTChart());
			}
			FileOutputStream fileOut = new FileOutputStream(new File(fileName + "Sprint-StoryTracker-Template1.xlsx"));
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	private static Map<String, String> readFile(String name) {
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

}
