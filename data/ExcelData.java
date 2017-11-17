package com.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.PreparedStatement;
import java.sql.SQLException;
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

public class ExcelData {
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

	/**
	 * 递归读取文件
	 */
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

	/**
	 * 石油化工类
	 */

	public static void insertChemical(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				String departName = getCellValue(row.getCell(1));
				if (StringUtils.isEmpty(departName)) {
					continue;
				}
				String departAddress = getCellValue(row.getCell(2));
				String fireTeamName = getCellValue(row.getCell(3));
				String productType = getCellValue(row.getCell(4));
				String maxStorage = getCellValue(row.getCell(5));
				String eastLongitudeX = getCellValue(row.getCell(6));
				String northLongitudeY = getCellValue(row.getCell(7));
				String departType = getCellValue(row.getCell(8));
				String managerName = getCellValue(row.getCell(9));
				String phoneNum = getCellValue(row.getCell(10));
				try {
					if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
						eastLongitudeX = getCellValue(row.getCell(7));
						northLongitudeY = getCellValue(row.getCell(6));
					}
				} catch (Exception e) {
					System.out.println("error路径--->" + file.getAbsolutePath());
					System.out.println("error路径--->" + e);
					e.printStackTrace();
				}

				String sql = "INSERT INTO \"t_chemical_dangerInfo\"(\"departName\",\"departAddress\","
						+ "\"fireTeamName\",\"productType\",\"maxStorage\",\"departType\","
						+ "\"managerName\",\"phoneNum\",\"provinceName\",\"eastLongitudeX\","
						+ "\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"departName\",\"provinceName\",\"departAddress\",\"fireTeamName\",\"managerName\",\"productType\",\"maxStorage\")DO update set \"departType\"=?,\"phoneNum\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

				try {
					stmt = Util.conn.prepareStatement(sql);
					stmt.setString(1, departName);
					stmt.setString(2, departAddress);
					stmt.setString(3, fireTeamName);
					stmt.setString(4, productType);
					stmt.setString(5, maxStorage);
					stmt.setString(6, departType);
					stmt.setString(7, managerName);
					stmt.setString(8, phoneNum);
					stmt.setString(9, province);
					stmt.setString(10, eastLongitudeX);
					stmt.setString(11, northLongitudeY);
					stmt.setString(12, departType);
					stmt.setString(13, phoneNum);
					stmt.setString(14, eastLongitudeX);
					stmt.setString(15, northLongitudeY);
					stmt.execute();
				} catch (SQLException e) {
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

	/**
	 * 水电站
	 */

	public static void insertWater(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String stationName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(stationName)) {
				continue;
			}
			String address = getCellValue(row.getCell(2));
			String mainDepartName = getCellValue(row.getCell(3));
			String managerName = getCellValue(row.getCell(4));
			String phoneNum = getCellValue(row.getCell(5));
			String totalStorage = getCellValue(row.getCell(6));
			String normalStorage = getCellValue(row.getCell(7));
			String installedCapacity = getCellValue(row.getCell(8));
			String rescueSituation = getCellValue(row.getCell(9));
			String eastLongitudeX = getCellValue(row.getCell(10));
			String northLongitudeY = getCellValue(row.getCell(11));
			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(11));
					northLongitudeY = getCellValue(row.getCell(10));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			String sql = "INSERT INTO \"t_water_electricStation\"(\"stationName\",\"address\","
					+ "\"mainDepartName\",\"managerName\",\"phoneNum\",\"totalStorage\","
					+ "\"normalStorage\",\"installedCapacity\",\"rescueSituation\",\"provinceName\",\"eastLongitudeX\","
					+ "\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"stationName\",\"provinceName\")DO update set \"address\"=?,\"mainDepartName\"=?,\"managerName\"=?,\"phoneNum\"=?,\"totalStorage\"=?,\"normalStorage\"=?,\"installedCapacity\"=?,\"rescueSituation\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, stationName);
				stmt.setString(2, address);
				stmt.setString(3, mainDepartName);
				stmt.setString(4, managerName);
				stmt.setString(5, phoneNum);
				stmt.setString(6, totalStorage);
				stmt.setString(7, normalStorage);
				stmt.setString(8, installedCapacity);
				stmt.setString(9, rescueSituation);
				stmt.setString(10, province);
				stmt.setString(11, eastLongitudeX);
				stmt.setString(12, northLongitudeY);

				stmt.setString(13, address);
				stmt.setString(14, mainDepartName);
				stmt.setString(15, managerName);
				stmt.setString(16, phoneNum);
				stmt.setString(17, totalStorage);
				stmt.setString(18, normalStorage);
				stmt.setString(19, installedCapacity);
				stmt.setString(20, rescueSituation);
				stmt.setString(21, eastLongitudeX);
				stmt.setString(22, northLongitudeY);
				stmt.execute();
			} catch (SQLException e) {
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

			// }
		}

	}

	/**
	 * 城市古建筑
	 */

	public static void insertCity(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String buildingName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(buildingName)) {
				continue;
			}
			String address = getCellValue(row.getCell(2));
			String buildYear = getCellValue(row.getCell(3));
			String buildingHeight = getCellValue(row.getCell(4));
			String upNumber = getCellValue(row.getCell(5));
			String downNumber = getCellValue(row.getCell(6));
			String maxLayerArea = getCellValue(row.getCell(7));
			String buildingArea = getCellValue(row.getCell(8));
			String buildingStructure = getCellValue(row.getCell(9));
			String buildingNature = getCellValue(row.getCell(10));
			String buildingType = getCellValue(row.getCell(11));
			String enterDepartName = getCellValue(row.getCell(12));
			String fireDepartName = getCellValue(row.getCell(13));
			String eastLongitudeX = getCellValue(row.getCell(14));
			String northLongitudeY = getCellValue(row.getCell(15));
			String neighborBuildingName = getCellValue(row.getCell(16));
			String facilitySituation = getCellValue(row.getCell(17));
			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(15));
					northLongitudeY = getCellValue(row.getCell(14));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			/*
			 * String sql =
			 * "INSERT INTO \"t_city_building\"(\"buildingName\",\"address\"," +
			 * "\"buildYear\",\"buildingHeight\",\"upNumber\",\"downNumber\"," +
			 * "\"maxLayerArea\",\"buildingArea\",\"buildingStructure\",\"buildingNature\","
			 * +
			 * "\"buildingType\",\"enterDepartName\",\"fireDepartName\",\"neighborBuildingName\",\"facilitySituation\",\"provinceName\",\"eastLongitudeX\",\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"buildingName\",\"provinceName\",\"address\")DO update set \"buildYear\"=?,\"buildingHeight\"=?,\"upNumber\"=?,\"downNumber\"=?,\"maxLayerArea\"=?,\"buildingArea\"=?,\"buildingStructure\"=?,\"buildingNature\"=?,\"buildingType\"=?,\"enterDepartName\"=?,\"fireDepartName\"=?,\"neighborBuildingName\"=?,\"facilitySituation\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?"
			 * ;
			 */
			String sql = "INSERT INTO \"t_city_building\"(\"buildingName\",\"address\","
					+ "\"buildYear\",\"buildingHeight\",\"upNumber\",\"downNumber\","
					+ "\"maxLayerArea\",\"buildingArea\",\"buildingStructure\",\"buildingNature\","
					+ "\"buildingType\",\"enterDepartName\",\"fireDepartName\",\"neighborBuildingName\",\"facilitySituation\",\"provinceName\",\"eastLongitudeX\",\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"buildingName\",\"provinceName\")DO update set \"address\"=?,\"buildYear\"=?,\"buildingHeight\"=?,\"upNumber\"=?,\"downNumber\"=?,\"maxLayerArea\"=?,\"buildingArea\"=?,\"buildingStructure\"=?,\"buildingNature\"=?,\"buildingType\"=?,\"enterDepartName\"=?,\"fireDepartName\"=?,\"neighborBuildingName\"=?,\"facilitySituation\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, buildingName);
				stmt.setString(2, address);
				stmt.setString(3, buildYear);
				stmt.setString(4, buildingHeight);
				stmt.setString(5, upNumber);
				stmt.setString(6, downNumber);
				stmt.setString(7, maxLayerArea);
				stmt.setString(8, buildingArea);
				stmt.setString(9, buildingStructure);
				stmt.setString(10, buildingNature);
				stmt.setString(11, buildingType);
				stmt.setString(12, enterDepartName);
				stmt.setString(13, fireDepartName);
				stmt.setString(14, neighborBuildingName);
				stmt.setString(15, facilitySituation);
				stmt.setString(16, province);
				stmt.setString(17, eastLongitudeX);
				stmt.setString(18, northLongitudeY);

				stmt.setString(19, address);
				stmt.setString(20, buildYear);
				stmt.setString(21, buildingHeight);
				stmt.setString(22, upNumber);
				stmt.setString(23, downNumber);
				stmt.setString(24, maxLayerArea);
				stmt.setString(25, buildingArea);
				stmt.setString(26, buildingStructure);
				stmt.setString(27, buildingNature);
				stmt.setString(28, buildingType);
				stmt.setString(29, enterDepartName);
				stmt.setString(30, fireDepartName);
				stmt.setString(31, neighborBuildingName);
				stmt.setString(32, facilitySituation);
				stmt.setString(33, eastLongitudeX);
				stmt.setString(34, northLongitudeY);
				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

			}
		}

	}

	/**
	 * 现役消防机构
	 */

	public static void insertFireStation(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String departName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(departName)) {
				continue;
			}
			String departType = getCellValue(row.getCell(2));
			String address = getCellValue(row.getCell(3));
			String eastLongitudeX = getCellValue(row.getCell(4));
			String northLongitudeY = getCellValue(row.getCell(5));
			String isCompileUnit = getCellValue(row.getCell(6));
			String isDutyUnit = getCellValue(row.getCell(7));
			String compileLeaderNum = getCellValue(row.getCell(8));
			String compileSoldierNum = getCellValue(row.getCell(9));
			String actualLeaderNum = getCellValue(row.getCell(10));
			String actualSoldierNum = getCellValue(row.getCell(11));

			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(5));
					northLongitudeY = getCellValue(row.getCell(4));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			String sql = "INSERT INTO \"t_fire_departInfo\"(\"departName\",\"departType\","
					+ "\"address\",\"isCompileUnit\",\"isDutyUnit\",\"compileLeaderNum\","
					+ "\"compileSoldierNum\",\"actualLeaderNum\",\"actualSoldierNum\",\"provinceName\","
					+ "\"eastLongitudeX\",\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"departName\",\"provinceName\")DO update set \"address\"=?,\"departType\"=?,\"isCompileUnit\"=?,\"isDutyUnit\"=?,\"compileLeaderNum\"=?,\"compileSoldierNum\"=?,\"actualLeaderNum\"=?,\"actualSoldierNum\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, departName);
				stmt.setString(2, departType);
				stmt.setString(3, address);
				stmt.setString(4, isCompileUnit);
				stmt.setString(5, isDutyUnit);
				stmt.setString(6, compileLeaderNum);
				stmt.setString(7, compileSoldierNum);
				stmt.setString(8, actualLeaderNum);
				stmt.setString(9, actualSoldierNum);
				stmt.setString(10, province);
				stmt.setString(11, eastLongitudeX);
				stmt.setString(12, northLongitudeY);
				stmt.setString(13, address);
				stmt.setString(14, departType);
				stmt.setString(15, isCompileUnit);
				stmt.setString(16, isDutyUnit);
				stmt.setString(17, compileLeaderNum);
				stmt.setString(18, compileSoldierNum);
				stmt.setString(19, actualLeaderNum);
				stmt.setString(20, actualSoldierNum);
				stmt.setString(21, eastLongitudeX);
				stmt.setString(22, northLongitudeY);
				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

			}
		}

	}

	/**
	 * 联勤保障单位
	 */

	public static void insertJoin(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String name = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(name)) {
				continue;
			}
			String address = getCellValue(row.getCell(2));
			String lng = getCellValue(row.getCell(3));
			String lat = getCellValue(row.getCell(4));
			String linkman = getCellValue(row.getCell(5));
			String phone = getCellValue(row.getCell(6));
			String type = getCellValue(row.getCell(7));
			String ability = getCellValue(row.getCell(8));

			try {
				if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
					lng = getCellValue(row.getCell(4));
					lat = getCellValue(row.getCell(3));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			String sql = "INSERT INTO \"t_joint_logistics_unit\"(\"name\",\"address\","
					+ "\"lng\",\"lat\",\"linkman\",\"phone\","
					+ "\"type\",\"ability\",\"province\")VALUES(?,?,?,?,?,?,?,?,?)ON conflict(\"name\",\"province\")DO update set \"address\"=?,\"ability\"=?,\"lng\"=?,\"lat\"=?,\"linkman\"=?,\"phone\"=?,\"type\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, name);
				stmt.setString(2, address);
				stmt.setString(3, lng);
				stmt.setString(4, lat);
				stmt.setString(5, linkman);
				stmt.setString(6, phone);
				stmt.setString(7, type);
				stmt.setString(8, ability);
				stmt.setString(9, province);

				stmt.setString(10, address);
				stmt.setString(11, ability);
				stmt.setString(12, lng);
				stmt.setString(13, lat);
				stmt.setString(14, linkman);
				stmt.setString(15, phone);
				stmt.setString(16, type);
				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

			}
		}

	}

	/**
	 * 多种形式
	 */

	public static void insertManyForms(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String departName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(departName) || departName.contains("机构名称")) {
				continue;
			}
			String departType = getCellValue(row.getCell(2));
			String managerDepartName = getCellValue(row.getCell(3));
			String address = getCellValue(row.getCell(4));
			String eastLongitudeX = getCellValue(row.getCell(5));
			String northLongitudeY = getCellValue(row.getCell(6));
			String contactPerson = getCellValue(row.getCell(7));
			String phoneNum = getCellValue(row.getCell(8));
			String personNum = getCellValue(row.getCell(9));
			String fireCarTypeNum = getCellValue(row.getCell(10));
			String fireEquipment = getCellValue(row.getCell(11));

			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(6));
					northLongitudeY = getCellValue(row.getCell(5));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			String sql = "INSERT INTO \"t_manyForms_fireTeam\"(\"departName\",\"departType\","
					+ "\"managerDepartName\",\"address\",\"contactPerson\",\"phoneNum\","
					+ "\"personNum\",\"fireCarTypeNum\",\"fireEquipment\",\"provinceName\","
					+ "\"eastLongitudeX\",\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"departName\",\"provinceName\")DO update set \"departType\"=?, \"managerDepartName\"=?,\"address\"=?,\"contactPerson\"=?,\"phoneNum\"=?,\"personNum\"=?,\"fireCarTypeNum\"=?,\"fireEquipment\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, departName);
				stmt.setString(2, departType);
				stmt.setString(3, managerDepartName);
				stmt.setString(4, address);
				stmt.setString(5, contactPerson);
				stmt.setString(6, phoneNum);
				stmt.setString(7, personNum);
				stmt.setString(8, fireCarTypeNum);
				stmt.setString(9, fireEquipment);
				stmt.setString(10, province);
				stmt.setString(11, eastLongitudeX);
				stmt.setString(12, northLongitudeY);

				stmt.setString(13, departType);
				stmt.setString(14, managerDepartName);
				stmt.setString(15, address);
				stmt.setString(16, contactPerson);
				stmt.setString(17, phoneNum);
				stmt.setString(18, personNum);
				stmt.setString(19, fireCarTypeNum);
				stmt.setString(20, fireEquipment);
				stmt.setString(21, eastLongitudeX);
				stmt.setString(22, northLongitudeY);
				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

			}
		}
	}

	/**
	 * 社区微型消防站
	 */

	public static void insertCommunityMicroStation(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			if (row == null)
				continue;
			String departName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(departName)) {
				continue;
			}
			String departType = getCellValue(row.getCell(2));
			String managerDepartName = getCellValue(row.getCell(3));
			String address = getCellValue(row.getCell(4));
			String eastLongitudeX = getCellValue(row.getCell(5));
			String northLongitudeY = getCellValue(row.getCell(6));
			String contactPerson = getCellValue(row.getCell(7));
			String phoneNum = getCellValue(row.getCell(8));
			String personNum = getCellValue(row.getCell(9));
			String fireCarTypeNum = getCellValue(row.getCell(10));
			String fireEquipment = getCellValue(row.getCell(11));

			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(6));
					northLongitudeY = getCellValue(row.getCell(5));
				}
			} catch (Exception e) {
				//System.out.println("error路径--->" + file.getAbsolutePath());
				//System.out.println("提示--->" + e);
				e.printStackTrace();
			}

			String sql = "INSERT INTO \"t_community_microFireStation\"(\"departName\",\"departType\","
					+ "\"managerDepartName\",\"address\",\"contactPerson\",\"phoneNum\","
					+ "\"personNum\",\"fireCarTypeNum\",\"fireEquipment\",\"provinceName\","
					+ "\"eastLongitudeX\",\"northLongitudeY\")VALUES(?,?,?,?,?,?,?,?,?,?,?,?)ON conflict(\"departName\",\"provinceName\",\"managerDepartName\")DO update set \"address\"=?,\"departType\"=?,\"contactPerson\"=?,\"phoneNum\"=?,\"personNum\"=?,\"fireCarTypeNum\"=?,\"fireEquipment\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, departName);
				stmt.setString(2, departType);
				stmt.setString(3, managerDepartName);
				stmt.setString(4, address);
				stmt.setString(5, contactPerson);
				stmt.setString(6, phoneNum);
				stmt.setString(7, personNum);
				stmt.setString(8, fireCarTypeNum);
				stmt.setString(9, fireEquipment);
				stmt.setString(10, province);
				stmt.setString(11, eastLongitudeX);
				stmt.setString(12, northLongitudeY);

				stmt.setString(13, address);
				stmt.setString(14, departType);
				stmt.setString(15, contactPerson);
				stmt.setString(16, phoneNum);
				stmt.setString(17, personNum);
				stmt.setString(18, fireCarTypeNum);
				stmt.setString(19, fireEquipment);
				stmt.setString(20, eastLongitudeX);
				stmt.setString(21, northLongitudeY);
				stmt.execute();
				//System.out.println(departName);

			} catch (SQLException e) {
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

	/**
	 * 核电站
	 */

	public static void insertNuclear(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;
		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String nuclearName = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(nuclearName)) {
				continue;
			}
			String nuclearAddress = getCellValue(row.getCell(2));
			String fireDepart = getCellValue(row.getCell(3));
			String eastLongitudeX = getCellValue(row.getCell(4));
			String northLongitudeY = getCellValue(row.getCell(5));
			String managerName = getCellValue(row.getCell(6));
			String phoneNum = getCellValue(row.getCell(7));
			String nuclearDescribe = getCellValue(row.getCell(8));

			try {
				if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
					eastLongitudeX = getCellValue(row.getCell(5));
					northLongitudeY = getCellValue(row.getCell(4));
				}
			} catch (Exception e) {
				System.out.println("error路径--->" + file.getAbsolutePath());
				System.out.println("提示--->" + e);
				e.printStackTrace();
			}
			String sql = "INSERT INTO \"t_nuclear_info\"(\"nuclearName\",\"nuclearAddress\","
					+ "\"fireDepart\",\"managerName\",\"phoneNum\",\"nuclearDescribe\","
					+ "\"provinceName\",\"eastLongitudeX\",\"northLongitudeY\""
					+ ")VALUES(?,?,?,?,?,?,?,?,?)ON conflict(\"nuclearName\",\"provinceName\")DO update set \"nuclearAddress\"=?,\"fireDepart\"=?,\"managerName\"=?,\"phoneNum\"=?,\"nuclearDescribe\"=?,\"eastLongitudeX\"=?,\"northLongitudeY\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, nuclearName);
				stmt.setString(2, nuclearAddress);
				stmt.setString(3, fireDepart);
				stmt.setString(4, managerName);
				stmt.setString(5, phoneNum);
				stmt.setString(6, nuclearDescribe);
				stmt.setString(7, province);
				stmt.setString(8, eastLongitudeX);
				stmt.setString(9, northLongitudeY);

				stmt.setString(10, nuclearAddress);
				stmt.setString(11, fireDepart);
				stmt.setString(12, managerName);
				stmt.setString(13, phoneNum);
				stmt.setString(14, nuclearDescribe);
				stmt.setString(15, eastLongitudeX);
				stmt.setString(16, northLongitudeY);
				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

			}
		}

	}

	/**
	 * 地震带
	 */

	public static void insertEarthquake(String province, File file, Sheet sheet, int rowNum, int xh) {
		PreparedStatement stmt = null;

		for (int i = xh + 1; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			// if (row != null) {
			if (row == null)
				continue;
			String name = getCellValue(row.getCell(1));
			if (StringUtils.isEmpty(name)) {
				continue;
			}
			String distributeArea = getCellValue(row.getCell(2));
			String rescueSituation = getCellValue(row.getCell(3));

			String sql = "INSERT INTO \"t_earthquake_info\"(\"name\",\"distributeArea\","
					+ "\"rescueSituation\",\"provinceName\""
					+ ")VALUES(?,?,?,?)ON conflict(\"name\",\"provinceName\")DO update set \"distributeArea\"=?,\"rescueSituation\"=?";

			try {
				stmt = Util.conn.prepareStatement(sql);
				stmt.setString(1, name);
				stmt.setString(2, distributeArea);
				stmt.setString(3, rescueSituation);
				stmt.setString(4, province);
				stmt.setString(5, distributeArea);
				stmt.setString(6, rescueSituation);

				stmt.execute();
			} catch (SQLException e) {
				e.printStackTrace();
			} finally {

				if (stmt != null) {
					try {
						stmt.close();
					} catch (SQLException e) {
						e.printStackTrace();
					}
				}

				// }

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
		int xh = 0;
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
				for (int k = 0; k <= rowNum; k++) {
					Row row2 = sheet.getRow(k);
					String res = getCellValue(row2.getCell(0));
					if (res.contains("序号")) {
						xh = k;
						// System.out.println("序号所在的行： " + xh);
						break;
					}
				}
				String title = cell.getStringCellValue();
				System.out.println(title);
				if (title.contains("石油化工")) {

				} else if (title.contains("水电站")) {
					 //insertWater(province, file, sheet, rowNum, xh);
				} else if (title.contains("城市综合体") || title.contains("文物古建筑")) {
					// insertCity(province, file, sheet, rowNum, xh);
				} else if (title.contains("现役消防机构")) {
					// insertFireStation(province, file, sheet, rowNum, xh);
				} else if (title.contains("联勤保障")) {
					// insertJoin(province, file, sheet, rowNum, xh);
				} else if (title.contains("多种形式消防队伍")) {
					//insertManyForms(province, file, sheet, rowNum, xh);
				} else if (title.contains("社区微型消防站") && !title.contains("队员")) {
					insertCommunityMicroStation(province, file, sheet, rowNum, xh);
				} else if (title.contains("核电站")) {
					// insertNuclear(province, file, sheet, rowNum, xh);
				} else if (title.contains("地震带")) {
					// insertEarthquake(province, file, sheet, rowNum, xh);
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
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09.13.1821");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\2017.08.30获取采集信息");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集0914.18.45");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\091514");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09151950");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09160900");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09160900差异文件");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09162017");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09181424");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集0919.10.57");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09191849");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09200904");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09201110");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09201403");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09201710");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09211000");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09211350");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09211857");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09220945");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09221330");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09221835");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09230920");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09231330");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09231900");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09251200");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09261127");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09261905");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09270845");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09271320");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09281327");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09281755");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\09291737");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\实战指挥平台数据采集09291639");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10011554");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10081332");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10131352");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10141910");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10161440");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10171018");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10171411");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10171929");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10172047");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10181820");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10182006");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10190900");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10191330");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\10191900");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10151900");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10202020");
		// pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10210915");
	//	pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10211850");
		//pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10221720");
		//pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10231335");
		//pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10241525");
		pathList.add("D:\\北京项目\\数据库表\\原始数据\\二次上报\\10250915");

		// pathList.add("D:\\北京项目\\test");

		try {
			System.setOut(new PrintStream(new FileOutputStream("D:\\log.txt")));
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
