package com.data.poi.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	public static Map<String, Set<String>> map = new HashMap<String, Set<String>>();
	public static Map<String, Set<String>> map2 = new HashMap<String, Set<String>>();

	public static String getCellValue(Cell cell) {
		// 判断是否为null或空串
		if (cell == null || cell.toString().trim().equals("")) {
			return "";
		}
		String cellValue = "";

		int cellType = cell.getCellType();

		switch (cellType) {
		case Cell.CELL_TYPE_STRING:
			cellValue = cell.getStringCellValue().trim();
			cellValue = StringUtils.isEmpty(cellValue) ? "" : cellValue;
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) { // 判断日期类型
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
				cellValue = df.format(cell.getDateCellValue());
			} else {
				cellValue = new DecimalFormat("#.##########").format(cell.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;

		default:
			cellValue = "";
			break;
		}

		return cellValue.trim();
	}

	private static void readFile(File file, String province) {

		if (file.exists()) {
			if (file.isDirectory()) {
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

	public static Workbook loadExcel(File file) {
		InputStream input = null;
		String fileName = file.getName();
		Workbook wb = null;
		try {
			input = new FileInputStream(file);

			try {
				if (fileName.endsWith(".xls")) {
					wb = new HSSFWorkbook(input); // .xls文件
				} else {
					wb = new XSSFWorkbook(input);// .xlsx文件
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return wb;
	}

	public static void readExcel(File file, String province) {
		try {
			Workbook wb = loadExcel(file);
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
				// 第一行
				Row row = sheet.getRow(0);
				if (row == null) {
					continue;
				}
				// 获取每个文档的头部标题，第一行
				String title = getCellValue(row.getCell(0));
				save(province, title);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void getMap(String province, String str) {

		if (map.containsKey(province)) {
			map.get(province).add(str);
		} else {
			Set<String> set = new HashSet<String>();
			set.add(str);
			map.put(province, set);
		}

	}

	public static void save(String province, String title) {
		if (title.contains("城市综合体") || title.contains("文物古建筑")) {
			getMap(province, Constant.str1);
		} else if (title.contains("石油化工")) {
			getMap(province, Constant.str2);
		} else if (title.contains("地震带")) {
			getMap(province, Constant.str3);
		} else if (title.contains("核电站")) {
			getMap(province, Constant.str4);
		} else if (title.contains("水电站")) {
			getMap(province, Constant.str5);
		} else if (title.contains("现役消防机构")) {
			getMap(province, Constant.str6);
		} else if (title.contains("多种形式消防队伍信息")) {
			getMap(province, Constant.str8);
		} else if (title.contains("社区微型消防站") && !title.contains("队员")) {
			getMap(province, Constant.str7);
		} else if (title.contains("执勤人员")) {
			getMap(province, Constant.str9);
		} else if (title.contains("灭火药剂")) {
			getMap(province, Constant.str10);
		} else if (title.contains("后勤保障")) {
			getMap(province, Constant.str11);
		} else if (title.contains("通信保障")) {
			getMap(province, Constant.str12);
		} else if (title.contains("应急联动")) {
			getMap(province, Constant.str16);
		} else if (title.contains("灭火救援专家")) {
			getMap(province, Constant.str15);
		} else if (title.contains("联勤保障")) {
			getMap(province, Constant.str14);
		} else if (title.contains("特种装备")) {
			specialEquipment();
		}
	}

	/**
	 * 特种装备信息的判断
	 */
	public static void specialEquipment() {
		File file = new File("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集\\全国各省装备信息汇总\\特种装备基础数据汇总.xls");
		Workbook wb = loadExcel(file);
		int sheetNumber = wb.getNumberOfSheets();
		// 循环每一sheet，并处理当前循环页
		for (int i = 0; i < sheetNumber; i++) {
			Sheet sheet = wb.getSheetAt(i);
			// 判断sheet是否隐藏
			if (sheet == null || wb.isSheetHidden(i)) {
				continue;
			}
			// 处理当前页，循环读取每一行
			int rowNum = sheet.getLastRowNum();// 行数
			for (int j = 4; j <= rowNum; j++) {
				Row row = sheet.getRow(j);
				String pro = getCellValue(row.getCell(6));
				if (pro.contains("黑龙江") || pro.contains("内蒙古")) {
					pro = pro.substring(0, 3);
				} else
					pro = pro.substring(0, 2);

				getMap(pro, Constant.str13);

			}

		}
	}

	public static void sysPrint() {
		StringBuffer buffer = new StringBuffer();
		String line = System.getProperty("line.separator");
		Set<String> keySet = map.keySet();
		for (String key : keySet) {
			Set<String> valueSet = map.get(key);
			for (String value : valueSet) {
				System.out.println(key + " -->  " + value);
				buffer.append(key + " -->  " + value).append(line);
			}
		}
		try {

			FileWriter fw = new FileWriter("D:\\test.txt");
			fw.write(buffer.toString());
			fw.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	@SuppressWarnings("null")
	public static void writeExcel(String path) {
		File file = new File(path);
		OutputStream out = null;
		try {
			Workbook wb = loadExcel(file);
			// 获取sheet的个数
			int sheetNumber = wb.getNumberOfSheets();
			// 循环每一sheet，并处理当前循环页
			for (int i = 0; i < sheetNumber; i++) {
				Sheet sheet = wb.getSheetAt(i);
				if (sheet == null || wb.isSheetHidden(i)) {
					continue;
				}

				int rowNum = sheet.getLastRowNum();// 行数
				Row row = sheet.getRow(0);
				int colNum = row.getPhysicalNumberOfCells();// 列数
				// 样式
				CellStyle style1 = wb.createCellStyle();
				CellStyle style = wb.createCellStyle();
				Font font = wb.createFont();

				for (int r = 5; r < rowNum - 1; r++) {
					Row row2 = sheet.getRow(r);
					// 获取到每个省份
					String province = getCellValue(row2.getCell(0));
					System.out.println("province: " + province);
					Set<String> set = map2.get(province);
					Set<Integer> res = new HashSet<Integer>();
					for (String value : set) {
						int temp = Integer.parseInt(value);
						res.add(temp);
					}
					// 获取全部分类
					for (int j = 1; j < colNum; j++) {
						Cell cell = row2.getCell(j);
						if (res.contains(j)) {
							cell.setCellValue("√");
							font.setFontName("黑体");
							font.setFontHeightInPoints((short) 10);// 设置字体大小
							style1.setFont(font);
							style1.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 下边框
							style1.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
							style1.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
							style1.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框
							style1.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平居中
							cell.setCellStyle(style1);
						} else {
							cell.setCellValue("×");
							font.setFontName("黑体");
							font.setFontHeightInPoints((short) 10);// 设置字体大小
							style.setFont(font);
							style.setFillForegroundColor(HSSFColor.GOLD.index);// 背景色
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
							style.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 下边框
							style.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
							style.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
							style.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框
							style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平居中
							cell.setCellStyle(style);
						}
					}
				}
				SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
				String time = sdf.format(new Date());
				File newFile = new File("D:\\" + "各总队基础数据采集情况统计表" + time + ".xls");
				out = new FileOutputStream(newFile);
				wb.write(out);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out == null) {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

	}

	/**
	 * 将map中保存的类型转换为对应的数字
	 */

	public static void trans() {
		Set<String> keySet = map.keySet();

		for (String key : keySet) {
			Set<String> valueSet = map.get(key);
			Set<String> set = new HashSet<String>();
			for (String value : valueSet) {
				if (value.contains("城市综合体") || value.contains("文物古建筑")) {
					set.add("1");
				} else if (value.contains("石油化工")) {
					set.add("2");
				} else if (value.contains("地震带")) {
					set.add("3");
				} else if (value.contains("核电站")) {
					set.add("4");
				} else if (value.contains("水电站")) {
					set.add("5");
				} else if (value.contains("现役消防机构")) {
					set.add("6");
				} else if (value.contains("微型消防站")) {
					set.add("7");
				} else if (value.contains("企业专职")) {
					set.add("8");
				} else if (value.contains("执勤人员")) {
					set.add("9");
				} else if (value.contains("灭火药剂")) {
					set.add("10");
				} else if (value.contains("后勤保障")) {
					set.add("11");
				} else if (value.contains("通信保障")) {
					set.add("12");
				} else if (value.contains("特种装备")) {
					set.add("13");
				} else if (value.contains("联勤保障")) {
					set.add("14");
				} else if (value.contains("灭火救援专家")) {
					set.add("15");
				} else if (value.contains("应急联动")) {
					set.add("16");
				}

			}
			map2.put(key, set);
		}
	}

	public static void main(String[] args) {

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
		for (int i = 0; i < pathList.size(); i++) {
			File file = new File(pathList.get(i));
			File[] files = file.listFiles();
			System.out.println("正在读取文件。。。");
			for (int j = 0; j < files.length; j++) {
				readFile(files[j], null);
			}
		}
		trans();
		String path = "D:\\模板.xls";
		writeExcel(path);
		//
		// sysPrint();
		// SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
		// System.out.println(sdf.format(new Date()));
		System.out.println("结束。。。");

	}

}
