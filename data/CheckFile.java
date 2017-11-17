package com.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CheckFile {
	private static FormulaEvaluator evaluator;

	/**
	 * 判断单元格数据类型
	 */
	public static String getCellValue(Cell cell) {
		// 判断是否为null或空串
		if (cell == null || cell.toString().trim().equals("")) {
			return "";
		}
		String cellValue = "";
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) { // 表达式类型
			cellType = evaluator.evaluate(cell).getCellType();
		}
		switch (cellType) {
		case Cell.CELL_TYPE_STRING:
			cellValue = cell.getStringCellValue().trim();
			cellValue = StringUtils.isEmpty(cellValue) ? "" : cellValue;
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) { // 判断日期类型
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
				cellValue = df.format(cell.getDateCellValue());
			} else {// 判断数据格式，保留10位小数
				cellValue = new DecimalFormat("#.##########").format(cell.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:// 布尔型
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		default:
			cellValue = "";
			break;
		}
		return cellValue.trim();
	}

	/**
	 * 递归读取文件
	 */
	private static void readFile(File file, String province) {
		if (file.exists()) {
			if (file.isDirectory()) {
				//判断省份，获取相应的文件夹名称
				province = province == null ? file.getName() : province;
				File[] files = file.listFiles();
				for (int i = 0; i < files.length; i++) {
					readFile(files[i], province);
				}
			} else {
				if (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) {
					readExcel(file, province);
				}
			}
		}
	}

	/**
	 * 读取Excel文件
	 * 
	 * @param file
	 * @param province
	 */
	public static void readExcel(File file, String province) {
		InputStream input = null;
		String fileName = file.getName();
		try {
			input = new FileInputStream(file);
			Workbook wb = null;
			try {
				if (fileName.endsWith(".xls")) {
					wb = new HSSFWorkbook(input); // .xls文件
				} else {
					wb = new XSSFWorkbook(input);// .xlsx文件
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			evaluator = wb.getCreationHelper().createFormulaEvaluator();
			// 获取sheet的个数
			int sheetNumber = wb.getNumberOfSheets();
			// 循环每一sheet，并处理当前循环页
			for (int i = 0; i < sheetNumber; i++) {
				Sheet sheet = wb.getSheetAt(i);
				// 判断sheet是否隐藏
				if (sheet == null || wb.isSheetHidden(i)) {
					continue;
				}
				// 处理当前页，循环读取每一行
				int rowNum = sheet.getLastRowNum();
				// 第一行
				Row row = sheet.getRow(0);
				if (row == null) {
					continue;
				}
				// 获取每个文档的头部标题，第一行
				Cell cell = row.getCell(0);
				if (cell == null) {
					continue;
				}
				//String title = cell.getStringCellValue();
				List<String> list = new ArrayList<String>();
				for (int k = 0; k <= rowNum; k++) {
					Row row2 = sheet.getRow(k);
					String res = getCellValue(row2.getCell(0));
					list.add(res);
					if (res.contains("序号"))
						break;
				}
				if (!(list.contains("序号"))) {
					System.out.println(" 无序号：  " + file.getAbsolutePath());
				}

			}
		} catch (Exception e) {
			System.out.println("异常文件： " + file.getAbsolutePath());
			System.out.println(e);
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		Util.getConnection();
		List<String> pathList = new ArrayList<String>();
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09.13.1821");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\2017.08.30获取采集信息");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集0914.18.45");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\091514");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09151950");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09160900");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09160900差异文件");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09162017");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\09181424");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集0919.10.57");
		// pathList.add("D:\\北京项目\\test");
		try {
			System.setOut(new PrintStream(new FileOutputStream("D:\\log2.txt")));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		for (int i = 0; i < pathList.size(); i++) {
			File file = new File(pathList.get(i));
			File[] files = file.listFiles();
			for (int j = 0; j < files.length; j++) {
				readFile(files[j], null);
			}
		}
	}

}
