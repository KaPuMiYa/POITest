package com.data.province;

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
import java.util.LinkedList;
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

public class SqlQuery2 {

	public static Map<String, LinkedList<String>> map = new HashMap<String, LinkedList<String>>();
	
	public static Map<String, LinkedList<String>> map2 = new HashMap<String, LinkedList<String>>();

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
						String nameRes = rs.getString("name");
						if (map2.containsKey(null)) {
							map2.remove(null);
						}
						if (map2.containsKey(nameRes)) {
							map2.get(nameRes).add(str);
						} else {
							LinkedList<String> al = new LinkedList<String>();
							if (str == null) {
								al.add("0");
							} else {
								al.add(str);
							}
							map2.put(nameRes, al);
						}
						System.out.println(nameRes + " -->  " + str);
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

	public static void writeExcel(String path) {
		File file = new File(path);
		OutputStream out = null;
		try {
			Workbook wb = loadExcel(file);
			Sheet sheet = wb.getSheetAt(0);
			int rowNum = sheet.getLastRowNum();
			int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
			Set<String> keySet = map2.keySet();

			for (int i = 3; i < rowNum; i++) {
				Row row = sheet.getRow(i);
				String res = getCellValue(row.getCell(0));
				for (String key : keySet) {
					LinkedList<String> valueSet = map2.get(key);
					if (res.equals(key)) {
						for (int j = 1; j < colNum; j++) {
							row.getCell(j).setCellValue(valueSet.get(j - 1));
						}
					}
				}
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

	public static void writeCell(String res, Row row, int k) {
		Set<String> keySet = map2.keySet();
		for (String key : keySet) {
			List<String> valueSet = map2.get(key);
			if (res.equals(key)) {
				row.getCell(k).setCellValue(valueSet.get(0));
				row.getCell(k + 1).setCellValue(valueSet.get(1));
			}

		}
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
		String path = "D:\\北京项目\\数据库表\\数据库统计汇总\\数据汇总统计分省份.xls";
		writeExcel(path); // 向模板文件中写入数据
		System.out.println("结束。。。");
	}

}
