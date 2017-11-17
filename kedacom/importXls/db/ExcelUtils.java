/**
 * 
 */
package com.kedacom.importXls.db;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
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

/**
 * @author hewenquan
 *
 */
public class ExcelUtils {

	/**
	 * 坐标表头部包含字段
	 */
	public static String COORDINATE = "坐标信息";

	public static String CHEMICAL = "石油化工";

	public static String MINI_FIRE_MEMBER = "社区微型消防站队员";

	public static String DUTY_PERSON = "执勤人员";

	public static String EXTINGUISHING = "灭火药剂";

	public static String SUPPORT_UNIT = "应急保障单位";

	public static String EMERGENCY_UNIT = "应急联动单位";

	public static String SPECIALIST = "灭火救援专家";

	public static String EQUIPMENT = "特种装备";
	public static String CITY = "城市综合体、文物古建筑信息";
	

	private static FormulaEvaluator evaluator;

	public static void jxlReadFile(File file, String province) {
		jxl.Workbook wb = null;
		try {
			wb = jxl.Workbook.getWorkbook(file);
			int sheetNum = wb.getNumberOfSheets();

			for (int i = 0; i < sheetNum; i++) {
				jxl.Sheet sheet = wb.getSheet(i);
				int rowNum = sheet.getRows();
				jxl.Cell[] cells = sheet.getRow(0);
				String title = cells[0].getContents();
				System.out.println(title);

				if (title.contains(ExcelUtils.COORDINATE)) {

					for (int start = 3; start < rowNum; start++) {
						jxl.Cell[] cs = sheet.getRow(start);
						if (cs != null && cs.length > 0) {

							if (cs.length == 4) {
								System.out.println(cells[1].getContents());
								System.out.println(cells[2].getContents());
								System.out.println(cells[3].getContents());

							} else if (cs.length == 5) {// 有地址时
								System.out.println(cells[1].getContents());
								System.out.println(cells[2].getContents());
								System.out.println(cells[3].getContents());
								System.out.println(cells[4].getContents());
							} else {// 有地址，类别时
								for (int j = 1; j < cs.length; j++) {
									System.out.println(cells[j].getContents());

								}
							}
						}

					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				wb.close();
			}
		}
	}

	public static void readFile(File file, String province) {

		FileInputStream in = null;
		String fileName = file.getName();
		try {
			in = new FileInputStream(file);
			Workbook wb = null;
			try {
				if (fileName.endsWith(".xls")) {
					wb = new HSSFWorkbook(in);// xls文件
				} else {
					wb = new XSSFWorkbook(in);// xlsx文件
				}
			} catch (Exception e) {
				e.printStackTrace();
				return;
			}
			evaluator = wb.getCreationHelper().createFormulaEvaluator();

			int sheetNum = wb.getNumberOfSheets();
			for (int i = 0; i < sheetNum; i++) {
				Sheet sheet = wb.getSheetAt(i);
				if (wb.isSheetHidden(i)) {
					return;
				}
				int rowNum = sheet.getLastRowNum();

				Row row = sheet.getRow(0);
				if (row == null) {
					System.out.println(province + ":" + fileName);
					return;
				}

				Cell cell = row.getCell(0);
				if (cell == null) {
					System.out.println(province + ":" + fileName);
					return;
				}
				String title = cell.getStringCellValue();
				System.out.println(title);
				// 消防队坐标坐标
				if (title.contains(ExcelUtils.COORDINATE)) {

					//importCoordante(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.CHEMICAL)) {// 石油化工
					//importChemical(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.MINI_FIRE_MEMBER)) {// 社区微型消防站队员信息
					//importMiniFireMember(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.DUTY_PERSON)) {// 执勤人员
					importDutyPerson(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.EXTINGUISHING)) {// 灭火药剂
					//importExtinguishing(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.SUPPORT_UNIT)) {// 应急保障
					//importSupportUnit(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.EMERGENCY_UNIT)) {// 应急联动保障单位
					//importEmergencyUnit(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.SPECIALIST)) {// 灭火专家
					//importSpecialist(file, province, sheet, rowNum);
				} else if (title.contains(ExcelUtils.EQUIPMENT)) {// 特殊装备
					//importEquipment(file, province, sheet, rowNum);
				}else if (title.contains(ExcelUtils.CITY)) {// 城市综合体、文物古建筑信息
					//importEquipment(file, province, sheet, rowNum);
				}
				
			}

		} catch (

		Exception e) {
			e.printStackTrace();
			System.out.println("异常文件：" + file.getAbsolutePath());
		} finally {

			if (in != null) {
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

	}

	/**
	 * 插入坐标信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importCoordante(File file, String province, Sheet sheet, int rowNum) throws SQLException {
		for (int start = 3; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				int cellNum = r.getLastCellNum();
				String name = "";
				String lng = "";
				String lat = "";
				String type = "";
				String address = "";
				name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}
				if (cellNum == 4) {
					lng = getCellValueByCell(r.getCell(2));
					lat = getCellValueByCell(r.getCell(3));

					try {
						if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
							lng = getCellValueByCell(r.getCell(3));
							lat = getCellValueByCell(r.getCell(2));
						}
					} catch (Exception e) {
						System.out.println(file.getAbsolutePath());
						e.printStackTrace();
					}

				} else if (cellNum == 5) {// 有地址时
					address = getCellValueByCell(r.getCell(2));
					lng = getCellValueByCell(r.getCell(3));
					lat = getCellValueByCell(r.getCell(4));

					try {
						if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
							lng = getCellValueByCell(r.getCell(4));
							lat = getCellValueByCell(r.getCell(3));
						}
					} catch (Exception e) {
						System.out.println(file.getAbsolutePath());
						e.printStackTrace();
					}
				} else {// 有地址，类别时
					type = getCellValueByCell(r.getCell(2));
					address = getCellValueByCell(r.getCell(3));
					lng = getCellValueByCell(r.getCell(4));
					lat = getCellValueByCell(r.getCell(5));

					try {
						if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
							lng = getCellValueByCell(r.getCell(5));
							lat = getCellValueByCell(r.getCell(4));
						}
					} catch (Exception e) {
						System.out.println(file.getAbsolutePath());
						e.printStackTrace();
					}
				}

				String sql = "insert into t_office_coordinate(name, lng, lat, type, address, province) values(?, ?, ?, ?, ?, ?) "
						+ "ON conflict(name,province) DO UPDATE SET lng = ?, lat = ?, type = ?, address = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, lng);
					statement.setString(3, lat);
					statement.setString(4, type);
					statement.setString(5, address);
					statement.setString(6, province);
					statement.setString(7, lng);
					statement.setString(8, lat);
					statement.setString(9, type);
					statement.setString(10, address);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}
				}

			}
		}
	}

	/**
	 * 插入石油化工、易燃易爆等化学危险品单位信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importChemical(File file, String province, Sheet sheet, int rowNum) throws SQLException {
		for (int start = 5; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String departName = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(departName)) {
					continue;
				}
				String departAddress = getCellValueByCell(r.getCell(2));
				String fireTeamName = getCellValueByCell(r.getCell(3));
				String productType = getCellValueByCell(r.getCell(4));
				String maxStorage = getCellValueByCell(r.getCell(5));

				String eastLongitudeX = getCellValueByCell(r.getCell(6));
				String northLongitudeY = getCellValueByCell(r.getCell(7));

				String departType = getCellValueByCell(r.getCell(8));
				String managerName = getCellValueByCell(r.getCell(9));
				String phoneNum = getCellValueByCell(r.getCell(10));

				try {
					if (Double.parseDouble(eastLongitudeX) < Double.parseDouble(northLongitudeY)) {
						eastLongitudeX = getCellValueByCell(r.getCell(7));
						northLongitudeY = getCellValueByCell(r.getCell(6));
					}
				} catch (Exception e) {
					System.out.println(file.getAbsolutePath());
					e.printStackTrace();
				}

				String sql = "insert into \"t_chemical_dangerInfo\"(\"departName\", \"departAddress\", \"fireTeamName\", \"productType\", \"maxStorage\", \"departType\",\"managerName\",\"phoneNum\",\"provinceName\",\"eastLongitudeX\",\"northLongitudeY\")"
						+ " values(?, ?, ?, ?, ?, ?,?, ?, ?, ?, ?)  ON conflict(\"departName\", \"departAddress\",\"provinceName\") DO UPDATE SET \"departAddress\" = ?, \"fireTeamName\" = ?, \"productType\" = ?, \"maxStorage\" = ?,\"departType\"=?,"
						+ "\"managerName\"=?,\"phoneNum\"=?, \"eastLongitudeX\"=?,\"northLongitudeY\" = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, departName);
					statement.setString(2, departAddress);
					statement.setString(3, fireTeamName);
					statement.setString(4, productType);
					statement.setString(5, maxStorage);
					statement.setString(6, departType);
					statement.setString(7, managerName);
					statement.setString(8, phoneNum);
					statement.setString(9, province);
					statement.setString(10, eastLongitudeX);
					statement.setString(11, northLongitudeY);
					statement.setString(12, departAddress);
					statement.setString(13, fireTeamName);
					statement.setString(14, productType);
					statement.setString(15, maxStorage);
					statement.setString(16, departType);
					statement.setString(17, managerName);
					statement.setString(18, phoneNum);
					statement.setString(19, eastLongitudeX);
					statement.setString(20, northLongitudeY);
					statement.execute();
					System.out.println(departName);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入微型消防站队员信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importMiniFireMember(File file, String province, Sheet sheet, int rowNum) throws SQLException {
		for (int start = 4; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}
				String idNumber = getCellValueByCell(r.getCell(2));
				String institution = getCellValueByCell(r.getCell(3));
				String type = getCellValueByCell(r.getCell(4));
				String sex = getCellValueByCell(r.getCell(5));
				String nation = getCellValueByCell(r.getCell(6));
				String nativePlace = getCellValueByCell(r.getCell(7));
				String post = getCellValueByCell(r.getCell(8));
				String phone = getCellValueByCell(r.getCell(9));

				String sql = "insert into \"t_mini_fire_station_member\"(\"name\", \"idNumber\", \"institution\" , \"type\", \"sex\", \"nation\", \"nativePlace\", \"post\",\"phone\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"name\",\"idNumber\",\"province\") DO UPDATE SET \"institution\" = ?, \"type\" = ?, \"sex\" = ?, \"nation\" = ?,\"nativePlace\"=?,"
						+ "\"post\"=?, \"phone\"=? ";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, idNumber);
					statement.setString(3, institution);
					statement.setString(4, type);
					statement.setString(5, sex);
					statement.setString(6, nation);
					statement.setString(7, nativePlace);
					statement.setString(8, post);
					statement.setString(9, phone);
					statement.setString(10, province);
					statement.setString(11, institution);
					statement.setString(12, type);
					statement.setString(13, sex);
					statement.setString(14, nation);
					statement.setString(15, nativePlace);
					statement.setString(16, post);
					statement.setString(17, phone);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入执勤人员信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importDutyPerson(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		String title = getCellValueByCell(sheet.getRow(0).getCell(0));
		String reportUnit = getCellValueByCell(sheet.getRow(2).getCell(0)).replaceAll("填报单位", "").replaceAll(":", "")
				.replaceAll("：", "");
		for (int start = 4; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}
				if(name.equals("797")){
					System.out.println("");
				}
				
				String category = "";
				if (title.contains("公安消防")) {
					category = "公安消防机构";
				} else {
					category = "政府专职消防队、企业专职消防队";
				}

				String idNumber = getCellValueByCell(r.getCell(2)).replaceAll(",", "");
				String institution = getCellValueByCell(r.getCell(3));
				String type = getCellValueByCell(r.getCell(4));
				String sex = getCellValueByCell(r.getCell(5));
				String nation = getCellValueByCell(r.getCell(6));
				String nativePlace = getCellValueByCell(r.getCell(7));
				String post = getCellValueByCell(r.getCell(8));
				String phone = getCellValueByCell(r.getCell(9));

				String sql = "insert into \"t_duty_persons\"(\"name\", \"idNumber\", \"institution\" , \"type\", \"sex\", \"nation\", \"nativePlace\", \"post\",\"phone\",\"category\",\"reportUnit\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"name\",\"idNumber\",\"province\") DO UPDATE SET \"institution\" = ?, \"type\" = ?, \"sex\" = ?, \"nation\" = ?,\"nativePlace\"=?,"
						+ "\"post\"=?, \"phone\"=?, \"category\"= ?, \"reportUnit\"= ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, idNumber);
					statement.setString(3, institution);
					statement.setString(4, type);
					statement.setString(5, sex);
					statement.setString(6, nation);
					statement.setString(7, nativePlace);
					statement.setString(8, post);
					statement.setString(9, phone);
					statement.setString(10, category);
					statement.setString(11, reportUnit);
					statement.setString(12, province);
					statement.setString(13, institution);
					statement.setString(14, type);
					statement.setString(15, sex);
					statement.setString(16, nation);
					statement.setString(17, nativePlace);
					statement.setString(18, post);
					statement.setString(19, phone);
					statement.setString(20, category);
					statement.setString(21, reportUnit);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 灭火药剂信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importExtinguishing(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		for (int start = 4; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}

				String type = getCellValueByCell(r.getCell(2));
				String number = getCellValueByCell(r.getCell(3));
				String institution = getCellValueByCell(r.getCell(4));
				String date = getCellValueByCell(r.getCell(5));

				String sql = "insert into \"t_extinguishing_agent\"(\"name\", \"type\", \"number\" , \"institution\", \"date\", \"province\")"
						+ " values(?, ?, ?, ?, ?, ?)  ON conflict(\"name\",\"type\", \"institution\",\"province\") DO UPDATE SET \"number\" = ?, \"date\" = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, type);
					statement.setString(3, number);
					statement.setString(4, institution);
					statement.setString(5, date);
					statement.setString(6, province);
					statement.setString(7, number);
					statement.setString(8, date);

					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入应急保障单位信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importSupportUnit(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		String title = getCellValueByCell(sheet.getRow(0).getCell(0));
		for (int start = 5; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}

				String category = "";
				if (title.contains("后勤保障")) {
					category = "后勤保障";
				} else {
					category = "应急通信保障分队";
				}

				String address = getCellValueByCell(r.getCell(2));
				String lng = getCellValueByCell(r.getCell(3));
				String lat = getCellValueByCell(r.getCell(4));
				String linkman = getCellValueByCell(r.getCell(5));
				String phone = getCellValueByCell(r.getCell(6));
				String type = getCellValueByCell(r.getCell(7));
				String institution = getCellValueByCell(r.getCell(8));
				String ability = getCellValueByCell(r.getCell(9));
				String supportName = getCellValueByCell(r.getCell(10));
				try {
					if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
						lng = getCellValueByCell(r.getCell(4));
						lat = getCellValueByCell(r.getCell(3));
					}
				} catch (Exception e) {
					e.printStackTrace();
				}

				String sql = "insert into \"t_emergency_support_unit\"(\"name\", \"address\", \"lng\" , \"lat\", \"linkman\", \"phone\", \"type\", \"institution\",\"ability\",\"supportName\",\"category\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"name\",\"category\",\"province\") DO UPDATE SET \"address\" = ?, \"lng\" = ?, \"lat\" = ?, \"linkman\" = ?,\"phone\"=?,"
						+ "\"type\"=?, \"institution\"=?, \"ability\"= ?, \"supportName\"= ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, address);
					statement.setString(3, lng);
					statement.setString(4, lat);
					statement.setString(5, linkman);
					statement.setString(6, phone);
					statement.setString(7, type);
					statement.setString(8, institution);
					statement.setString(9, ability);
					statement.setString(10, supportName);
					statement.setString(11, category);
					statement.setString(12, province);
					statement.setString(13, address);
					statement.setString(14, lng);
					statement.setString(15, lat);
					statement.setString(16, linkman);
					statement.setString(17, phone);
					statement.setString(18, type);
					statement.setString(19, institution);
					statement.setString(20, ability);
					statement.setString(21, supportName);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入应急联动保障单位信息
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importEmergencyUnit(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		for (int start = 5; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}

				String address = getCellValueByCell(r.getCell(2));
				String lng = getCellValueByCell(r.getCell(3));
				String lat = getCellValueByCell(r.getCell(4));
				String director = getCellValueByCell(r.getCell(5));
				String phone = getCellValueByCell(r.getCell(6));
				String dutyTelephone = getCellValueByCell(r.getCell(7));
				String type = getCellValueByCell(r.getCell(8));
				try {
					if (Double.parseDouble(lng) < Double.parseDouble(lat)) {
						lng = getCellValueByCell(r.getCell(4));
						lat = getCellValueByCell(r.getCell(3));
					}
				} catch (Exception e) {
					System.out.println(file.getAbsolutePath());
					e.printStackTrace();
				}

				String sql = "insert into \"t_emergency_unit\"(\"name\", \"address\", \"lng\" , \"lat\", \"director\", \"phone\", \"dutyTelephone\",\"type\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"name\",\"province\") DO UPDATE SET \"address\" = ?, \"lng\" = ?, \"lat\" = ?, \"director\" = ?,\"phone\"=?,"
						+ "\"dutyTelephone\"=?, \"type\"=?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, address);
					statement.setString(3, lng);
					statement.setString(4, lat);
					statement.setString(5, director);
					statement.setString(6, phone);
					statement.setString(7, dutyTelephone);
					statement.setString(8, type);
					statement.setString(9, province);
					statement.setString(10, address);
					statement.setString(11, lng);
					statement.setString(12, lat);
					statement.setString(13, director);
					statement.setString(14, phone);
					statement.setString(15, dutyTelephone);
					statement.setString(16, type);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入灭火专家
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importSpecialist(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		for (int start = 5; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String name = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(name)) {
					continue;
				}

				String sex = getCellValueByCell(r.getCell(2));
				String education = getCellValueByCell(r.getCell(3));
				String post = getCellValueByCell(r.getCell(4));
				String idNumber = getCellValueByCell(r.getCell(5));
				String phone = getCellValueByCell(r.getCell(6));
				String institution = getCellValueByCell(r.getCell(7));
				String address = getCellValueByCell(r.getCell(8));
				String field = getCellValueByCell(r.getCell(9));
				String isArmy = getCellValueByCell(r.getCell(10));

				String sql = "insert into \"t_rescue_specialist\"(\"name\", \"sex\", \"education\" , \"post\", \"idNumber\", \"phone\", \"institution\", \"address\", \"field\", \"isArmy\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"name\", \"idNumber\",\"province\") DO UPDATE SET \"sex\" = ?, \"education\" = ?, \"post\" = ?, \"phone\" = ?, \"institution\" = ?,\"address\"=?, \"field\" = ?, \"isArmy\" = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, name);
					statement.setString(2, sex);
					statement.setString(3, education);
					statement.setString(4, post);
					statement.setString(5, idNumber);
					statement.setString(6, phone);
					statement.setString(7, institution);
					statement.setString(8, address);
					statement.setString(9, field);
					statement.setString(10, isArmy);
					statement.setString(11, province);
					statement.setString(12, sex);
					statement.setString(13, education);
					statement.setString(14, post);
					statement.setString(15, phone);
					statement.setString(16, institution);
					statement.setString(17, address);
					statement.setString(18, field);
					statement.setString(19, isArmy);
					statement.execute();
					System.out.println(name);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	/**
	 * 插入特殊装备
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importEquipment(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		boolean isAll = false;
		if (province.contains("全国各省")) {
			isAll = true;
		}

		for (int start = 5; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String type = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(type)) {
					continue;
				}

				String model = getCellValueByCell(r.getCell(2));
				String date = getCellValueByCell(r.getCell(3));
				String proform = getCellValueByCell(r.getCell(4));
				String number = getCellValueByCell(r.getCell(5));
				String institution = getCellValueByCell(r.getCell(6));
				String statDate = getCellValueByCell(r.getCell(7));

				if (isAll) {
					if (institution.contains("内蒙古")) {
						province = "内蒙古";
					} else if (institution.contains("黑龙江")) {
						province = "黑龙江";
					} else {
						province = institution.substring(0, 2);
					}
				}

				String sql = "insert into \"t_special_equipment\"(\"type\", \"model\", \"date\" , \"proform\", \"number\", \"institution\", \"statDate\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"type\", \"model\", \"institution\",\"province\") DO UPDATE SET \"date\" = ?, \"proform\" = ?, \"number\" = ?, \"statDate\" = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, type);
					statement.setString(2, model);
					statement.setString(3, date);
					statement.setString(4, proform);
					statement.setString(5, number);
					statement.setString(6, institution);
					statement.setString(7, statDate);
					statement.setString(8, province);
					statement.setString(9, date);
					statement.setString(10, proform);
					statement.setString(11, number);
					statement.setString(12, statDate);
					;
					statement.execute();
					System.out.println(type);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}
	
	/**
	 * 插入特殊装备
	 * 
	 * @param file
	 * @param province
	 * @param sheet
	 * @param rowNum
	 * @throws SQLException
	 */
	private static void importCity(File file, String province, Sheet sheet, int rowNum) throws SQLException {

		boolean isAll = false;
		if (province.contains("全国各省")) {
			isAll = true;
		}

		for (int start = 4; start < rowNum; start++) {
			Row r = sheet.getRow(start);
			if (r != null) {

				String type = getCellValueByCell(r.getCell(1));
				if (StringUtils.isBlank(type)) {
					continue;
				}

				String model = getCellValueByCell(r.getCell(2));
				String date = getCellValueByCell(r.getCell(3));
				String proform = getCellValueByCell(r.getCell(4));
				String number = getCellValueByCell(r.getCell(5));
				String institution = getCellValueByCell(r.getCell(6));
				String statDate = getCellValueByCell(r.getCell(7));

				if (isAll) {
					if (institution.contains("内蒙古")) {
						province = "内蒙古";
					} else if (institution.contains("黑龙江")) {
						province = "黑龙江";
					} else {
						province = institution.substring(0, 2);
					}
				}

				String sql = "insert into \"t_special_equipment\"(\"type\", \"model\", \"date\" , \"proform\", \"number\", \"institution\", \"statDate\",\"province\")"
						+ " values(?, ?, ?, ?, ?, ?, ?, ?)  ON conflict(\"type\", \"model\", \"institution\",\"province\") DO UPDATE SET \"date\" = ?, \"proform\" = ?, \"number\" = ?, \"statDate\" = ?";

				PreparedStatement statement = null;
				try {

					statement = DBUtils.connection.prepareStatement(sql);
					statement.setString(1, type);
					statement.setString(2, model);
					statement.setString(3, date);
					statement.setString(4, proform);
					statement.setString(5, number);
					statement.setString(6, institution);
					statement.setString(7, statDate);
					statement.setString(8, province);
					statement.setString(9, date);
					statement.setString(10, proform);
					statement.setString(11, number);
					statement.setString(12, statDate);
					;
					statement.execute();
					System.out.println(type);
				} catch (Exception e) {
					e.printStackTrace();
				} finally {

					if (statement != null) {
						statement.close();
					}

				}

			}
		}
	}

	// 获取单元格各类型值，返回字符串类型
	private static String getCellValueByCell(Cell cell) {
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
		case Cell.CELL_TYPE_STRING: // 字符串类型
			cellValue = cell.getStringCellValue().trim();
			cellValue = StringUtils.isEmpty(cellValue) ? "" : cellValue;
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔类型
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC: // 数值类型
			if (HSSFDateUtil.isCellDateFormatted(cell)) { // 判断日期类型
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
				cellValue = df.format(cell.getDateCellValue());
			} else { // 否
				cellValue = new DecimalFormat("#.##########").format(cell.getNumericCellValue());
			}
			break;
		default: // 其它类型，取空串吧
			cellValue = "";
			break;
		}
		return cellValue.trim();
	}

	public static void loadFile(File file, String province) {
		if (file.exists()) {
			if (file.isDirectory()) {
				province = province == null ? file.getName() : province;
				File[] files = file.listFiles();
				for (int i = 0; i < files.length; i++) {
					loadFile(files[i], province);
				}

			} else {
				if (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx"))
					ExcelUtils.readFile(file, province);
			}
		}
	}

	public static void main(String[] args) {

		DBUtils.connect();
		List<String> paths = new ArrayList<String>();
		paths.add("D:\\kedacom\\数据\\实战指挥平台数据采集");
		paths.add("D:\\kedacom\\数据\\2017.08.30获取采集信息");
		paths.add("D:\\kedacom\\数据\\实战指挥平台数据采集09.13.1821");

		for (int i = 0; i < paths.size(); i++) {
			File file = new File(paths.get(i));
			File[] files = file.listFiles();
			for (int j = 0; j < files.length; j++) {
				ExcelUtils.loadFile(files[j], null);
			}
		}
		DBUtils.Disconnect();

	}
}
