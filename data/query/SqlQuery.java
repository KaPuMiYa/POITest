package com.data.query;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SqlQuery {

	public static Map<String, List<String>> map = new HashMap<String, List<String>>();
	public static Map<String, List<String>> map2 = new HashMap<String, List<String>>();

	/**
	 * 创建sql集合
	 * 
	 */
	public static String city = "t_city_building";
	public static String water = "t_water_electricStation";
	public static String departName = "keda_xfjd_dwxx_clear";
	public static String departNameClean = "gis_xfjd_dwxx_clear";
	public static String buildingName = "a_fire_jzxx";
	public static String chemical = "t_chemical_dangerInfo";
	public static String nuclear = "t_nuclear_info";
	public static String earthquake = "a_fire_zpfa";
	public static String institutional = "institutional_data";
	public static String osmClean = "gis_osm_clean";
	public static String fireDepart = "t_fire_departInfo";
	public static String manyForms = "t_manyForms_fireTeam";
	public static String community = "t_community_microFireStation";
	public static String join = "t_joint_logistics_unit";
	public static String extinguish = "t_extinguishing_agent";
	public static String equipment = "t_special_equipment";
	public static String heightBuild = "buildingInfo";
	public static String durtyPerson = "t_duty_persons";
	public static String specialPerson = "t_rescue_specialist";

	public static void SqlList() {

		String name = "SELECT count(0) as num FROM \"";
		String temp = "WHERE \"eastLongitudeX\" ~ ? AND \"northLongitudeY\" ~ ?";
		String temp1 = "WHERE \"lng\" ~ ? AND \"lat\" ~ ?";
		String temp2 = "WHERE institution!=''";

		// 城市综合体
		String sql1 = name + city + "\"";
		String sql2 = sql1 + temp;
		putMapValue(city, sql1, sql2);
		// putMapValue(city, sql2);

		// 大型水电站
		String sql3 = name + water + "\"";
		String sql4 = sql3 + temp;
		putMapValue(water, sql3, sql4);

		// 单位信息
		String sql5 = name + departName + "\"";
		String sql6 = name + departNameClean + "\"";
		putMapValue(departNameClean, sql5, sql6);

		// 全部建筑
		String sql7 = name + buildingName + "\"";
		putMapValue(buildingName, sql7, sql7);

		// 石油化工
		String sql8 = name + chemical + "\"";
		String sql9 = sql8 + temp;
		putMapValue(chemical, sql8, sql9);

		// 核电站
		String sql10 = name + nuclear + "\"";
		String sql11 = sql10 + temp;
		putMapValue(nuclear, sql10, sql11);

		// 地震带
		String sql12 = name + earthquake + "\"";
		putMapValue(earthquake, sql12, sql12);

		// 全部消防机构
		String sql13 = name + institutional + "\"";
		String sql14 = name + osmClean + "\"";
		putMapValue(osmClean, sql13, sql14);

		// 现役消防队
		String sql15 = name + fireDepart + "\"";
		String sql16 = sql15 + temp;
		putMapValue(fireDepart, sql15, sql16);

		// 政府专职，企业专职
		String sql17 = name + manyForms + "\"";
		String sql18 = sql17 + temp;
		putMapValue(manyForms, sql17, sql18);

		// 微型消防站
		String sql19 = name + community + "\"";
		String sql20 = sql19 + temp;
		putMapValue(community, sql19, sql20);

		// 联勤保障
		String sql21 = name + join + "\"";
		String sql22 = sql21 + temp1;
		putMapValue(join, sql21, sql22);

		// 灭火药剂
		String sql23 = name + extinguish + "\"";
		String sql24 = sql23 + temp2;
		putMapValue(extinguish, sql23, sql24);
		// 特种装备
		String sql25 = name + equipment + "\"";
		String sql26 = sql25 + temp2;
		putMapValue(equipment, sql25, sql26);

		// 高层建筑
		String sql27 = name + heightBuild + "\"";
		putMapValue(heightBuild, sql27, sql27);
		// 执勤人员信息
		String sql28 = name + durtyPerson + "\"";
		putMapValue(durtyPerson, sql28, sql28);
		// 灭火救援专家
		String sql29 = name + specialPerson + "\"";
		putMapValue(specialPerson, sql29, sql29);

	}

	public static void putMapValue(String name, String sql1, String sql2) {
		if (map.containsKey(name)) {
			map.get(name).add(sql1);
			map.get(name).add(sql2);
		} else {
			List<String> list = new ArrayList<String>();
			list.add(sql1);
			list.add(sql2);
			map.put(name, list);
		}
	}

	public static void findData() {
		PreparedStatement stmt = null;
		String regx = "^[0-9]{2,3}.[0-9]{6,16}$";
		Set<String> keySet = map.keySet();
		for (String key : keySet) {
			List<String> list = map.get(key);
			for (String sql : list) {
				try {
					stmt = SQLConnect.conn.prepareStatement(sql);
					if (sql.contains("?")) {
						stmt.setString(1, regx);
						stmt.setString(2, regx);
					}
					ResultSet rs = stmt.executeQuery();
					while (rs.next()) {
						String str = rs.getString("num");
						if (map2.containsKey(key)) {
							map2.get(key).add(str);
						} else {
							List<String> al = new ArrayList<String>();
							al.add(str);
							map2.put(key, al);
						}
						System.out.println(key + " -->  " + str);
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
	@SuppressWarnings("null")
	public static void writeExcel(String path) {
		File file = new File(path);
		OutputStream out = null;
		try {
			Workbook wb = loadExcel(file);
			Sheet sheet = wb.getSheetAt(0);
			writeCell(sheet, 26, departNameClean);
			writeCell(sheet, 31, buildingName);
			writeCell(sheet, 32, heightBuild);
			writeCell(sheet, 35, city);
			writeCell(sheet, 36, chemical);
			writeCell(sheet, 38, nuclear);
			writeCell(sheet, 39, water);
			writeCell(sheet, 40, earthquake);
			writeCell(sheet, 41, osmClean);
			writeCell(sheet, 43, fireDepart);
			writeCell(sheet, 45, manyForms);
			writeCell(sheet, 46, community);
			writeCell(sheet, 54, join);
			writeCell(sheet, 57, extinguish);
			writeCell(sheet, 59, equipment);
			writeCell(sheet, 61, durtyPerson);
			writeCell(sheet, 62, specialPerson);

			System.out.println("写入完毕...。。。");

			// 针对模板文件另存，为文件名添加时间标志
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
			String time = sdf.format(new Date());
			File newFile = new File("D:\\北京项目\\数据库表\\数据库统计汇总\\" + "数据入库汇总报表" + time + ".xls");
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

	public static void writeCell(Sheet sheet, int index, String name) {
		Row row2 = sheet.getRow(index);
		List<String> list = map2.get(name);
		row2.getCell(5).setCellValue(list.get(0));
		row2.getCell(6).setCellValue(list.get(1));
		row2.getCell(7).setCellValue(name);

	}

	/**
	 * 打印map
	 */
	public static void sysPrint() {
		Set<String> keySet = map2.keySet();
		for (String key : keySet) {
			List<String> valueSet = map2.get(key);
			for (String value : valueSet) {
				System.out.println(key + " **********************  " + value);
			}
		}
	}

	public static void main(String[] args) {
		SQLConnect.getConnection();
		SqlList();
		findData();
		// sysPrint();
		String path = "D:\\北京项目\\数据库表\\数据库统计汇总\\数据入库汇总报表(模板).xls";
		writeExcel(path); // 向模板文件中写入数据
		System.out.println("结束。。。");
	}

}
