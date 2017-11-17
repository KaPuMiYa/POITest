package com.data.province;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class SqlQuery3 {

	public static Map<String, List<String>> map = new HashMap<String, List<String>>();
	public static Map<String, Map<String, List<String>>> map2 = new HashMap<String, Map<String, List<String>>>();
	/**
	 * 创建sql集合
	 * 
	 */
	public static String city = "t_city_building";
	public static String water = "t_water_electricStation";
	public static String chemical = "t_chemical_dangerInfo";
	public static String nuclear = "t_nuclear_info";
	public static String earthquake = "a_fire_zpfa";
	public static String fireDepart = "t_fire_departInfo";
	public static String manyForms = "t_manyForms_fireTeam";
	public static String community = "t_community_microFireStation";
	public static String join = "t_joint_logistics_unit";
	public static String extinguish = "t_extinguishing_agent";
	public static String equipment = "t_special_equipment";
	public static String durtyPerson = "t_duty_persons";
	public static String specialPerson = "t_rescue_specialist";

	public static void SqlList() {

		String name = "SELECT \"provinceName\" AS name,count(0) AS num FROM \"";
		String name2 = "SELECT \"province\" AS name,count(0) AS num FROM \"";

		String temp = "WHERE \"eastLongitudeX\" ~ ? AND \"northLongitudeY\" ~ ? GROUP BY \"provinceName\"";
		String temp1 = "WHERE \"lng\" ~ ? AND \"lat\" ~ ? GROUP BY \"province\"";
		String temp2 = "WHERE institution!='' GROUP BY \"province\"";

		// 城市综合体
		String sql1 = name + city + "\" GROUP BY \"provinceName\"";
		String sql2 = name + city + "\"" + temp;
		putMapValue(city, sql1, sql2);

		// 石油化工
		String sql3 = name + chemical + "\" GROUP BY \"provinceName\"";
		String sql4 = name + chemical + "\"" + temp;
		putMapValue(chemical, sql3, sql4);
		// 核电站
		String sql5 = name + nuclear + "\" GROUP BY \"provinceName\"";
		String sql6 = name + nuclear + "\"" + temp;
		putMapValue(nuclear, sql5, sql6);
		// 大型水库水电站
		String sql7 = name + water + "\" GROUP BY \"provinceName\"";
		String sql8 = name + water + "\"" + temp;
		putMapValue(water, sql7, sql8);
		// 现役
		String sql9 = name + fireDepart + "\" GROUP BY \"provinceName\"";
		String sql10 = name + fireDepart + "\"" + temp;
		putMapValue(fireDepart, sql9, sql10);
		// 政府企业专职
		String sql11 = name + manyForms + "\" GROUP BY \"provinceName\"";
		String sql12 = name + manyForms + "\"" + temp;
		putMapValue(manyForms, sql11, sql12);
		// 微型消防站
		String sql13 = name + community + "\" GROUP BY \"provinceName\"";
		String sql14 = name + community + "\"" + temp;
		putMapValue(community, sql13, sql14);
		// 联勤单位
		String sql15 = name2 + join + "\" GROUP BY \"province\"";
		String sql16 = name2 + join + "\"" + temp1;
		putMapValue(join, sql15, sql16);
		// 灭火药剂
		String sql17 = name2 + extinguish + "\" GROUP BY \"province\"";
		String sql18 = name2 + extinguish + "\"" + temp2;
		putMapValue(extinguish, sql17, sql18);

		// 特种装备
		String sql19 = name2 + equipment + "\" GROUP BY \"province\"";
		putMapValue(equipment, sql19, sql19);

		// 执勤人员
		String sql20 = name2 + durtyPerson + "\" GROUP BY \"province\"";
		putMapValue(durtyPerson, sql20, sql20);
		// 救援专家
		String sql22 = name2 + specialPerson + "\" GROUP BY \"province\"";
		putMapValue(specialPerson, sql22, sql22);

	}

	public static void putMapValue(String name, String sql1, String sql2) {
		if (map.containsKey(name)) {
			map.get(name).add(sql1);
			map.get(name).add(sql2);
		} else {
			LinkedList<String> list = new LinkedList<String>();
			list.add(sql1);
			list.add(sql2);
			map.put(name, list);
		}
	}

	public static void findData() {
		PreparedStatement stmt = null;
		String regx = "^[0-9]{2,3}.[0-9]{3,16}$";
		Set<String> keySet = map.keySet();
		for (String key : keySet) {
			List<String> list = map.get(key);
			Map<String, List<String>> tempMap = new HashMap<String, List<String>>();
			for (String sql : list) {
				try {
					stmt = SQLConnect.conn.prepareStatement(sql);
					if (sql.contains("?")) {
						stmt.setString(1, regx);
						stmt.setString(2, regx);
					}
					ResultSet rs = stmt.executeQuery();

					while (rs.next()) {
						String num = rs.getString("num");
						String nameRes = rs.getString("name");
						LinkedList<String> al = new LinkedList<String>();
						if (tempMap.containsKey(nameRes)) {
							tempMap.get(nameRes).add(num);
						} else {
							al.add(num);
							tempMap.put(nameRes, al);
						}
					}

				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					if (stmt != null) {
						try {
							stmt.close();
						} catch (SQLException e) {
							e.printStackTrace();
						}
					}
				}
			}

			map2.put(key, tempMap);

		}
		changeMap();

	}

	// 判断tempMap是否包含所有的省份，如果没有该省，则自动填充，补0
	public static void changeMap() {
		Set<String> keySet = map2.keySet();
		for (String key : keySet) {
			Map<String, List<String>> temp = map2.get(key);
			Set<String> proList = temp.keySet();

			String[] province = { "北京", "新疆", "重庆", "广东", "天津", "浙江", "广西", "内蒙古", "宁夏", "江西", "安徽", "贵州", "陕西", "辽宁",
					"山西", "青海", "四川", "江苏", "河北", "西藏", "福建", "吉林", "湖北", "海南", "上海", "云南", "甘肃", "湖南", "河南", "山东",
					"黑龙江" };
			List<String> list = Arrays.asList(province);
			for (String p : list) {
				if (proList.contains(p)) {
					List<String> li = temp.get(p);
					if (li.size() < 2) {
						li.add("0");
					}
				} else {
					List<String> li = new ArrayList();
					li.add("0");
					li.add("0");
					temp.put(p, li);
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
			if (fileName.endsWith(".xls")) {
				wb = new HSSFWorkbook(input); // .xls文件
			} else {
				wb = new XSSFWorkbook(input);// .xlsx文件
			}
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("异常文件：" + file.getAbsolutePath());
			System.out.println(e);
		}
		return wb;
	}

	/**
	 * 判断单元格数据类型，获取相应值
	 * 
	 * @param cell
	 * @return
	 */
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
			} else {// 判断数字格式，保留10位小数
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

	/**
	 * 向Excel模板中写入数据
	 * 
	 * @param path
	 */

	public static void writeExcel(String path) {
		File file = new File(path);
		OutputStream out = null;
		try {
			Workbook wb = loadExcel(file);
			Sheet sheet = wb.getSheetAt(0);
			int rowNum = sheet.getLastRowNum();
			for (int i = 3; i < rowNum; i++) {
				Row row = sheet.getRow(i);
				writeCell("t_city_building", row, 1);
				writeCell("t_chemical_dangerInfo", row, 3);
				writeCell("t_nuclear_info", row, 5);
				writeCell("t_water_electricStation", row, 7);
				writeCell("t_fire_departInfo", row, 9);
				writeCell("t_manyForms_fireTeam", row, 11);
				writeCell("t_community_microFireStation", row, 13);
				writeCell("t_joint_logistics_unit", row, 15);
				writeCell("t_extinguishing_agent", row, 17);
				writeCell("t_special_equipment", row, 19);
				writeCell("t_duty_persons", row, 21);
				writeCell("t_rescue_specialist", row, 23);
			}
			// 针对模板文件另存，为文件名添加时间标志
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
			String time = sdf.format(new Date());
			File newFile = new File("D:\\北京项目\\数据库表\\数据库统计汇总\\" + "数据汇总统计分省份" + time + ".xls");
			out = new FileOutputStream(newFile);
			wb.write(out);

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

	public static void writeCell(String s1, Row row, int k) {
		Set<String> keySet = map2.keySet();
		for (String key : keySet) {
			if (key.equals(s1)) {
				String province = getCellValue(row.getCell(0));
				Map<String, List<String>> temp = map2.get(key);
				Set<String> proList = temp.keySet();
				List<String> list = temp.get(province);
				for (String s : proList) {
					if (s.equals(province)) {
						row.getCell(k).setCellValue(list.get(0));
						row.getCell(k + 1).setCellValue(list.get(1));
					}
				}
			}

		}
	}

	/**
	 * 打印map
	 */
	public static void sysPrint2() {
		Set<String> keySet = map2.keySet();
		for (String key : keySet) {
			Map<String, List<String>> temp = map2.get(key);
			Set<String> valueSet = temp.keySet();
			for (String value : valueSet) {
				List<String> li = temp.get(value);
				 System.out.println(key + " ********************** " + value );
			}
		}
	}

	public static void main(String[] args) {
		SQLConnect.getConnection();
		SqlList();
		findData();
		sysPrint2();
		String path = "D:\\北京项目\\数据库表\\数据库统计汇总\\数据汇总统计分省份.xls";
		writeExcel(path); // 向模板文件中写入数据
		System.out.println("结束。。。");
	}

}
