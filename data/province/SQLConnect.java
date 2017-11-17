package com.data.province;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import com.data.Util;

public class SQLConnect {
	public static String driver = "org.postgresql.Driver";
	public static Connection conn = null;

//	public static String username = "postgres";
//	public static String password = "123456";
//	public static String url = "jdbc:postgresql://localhost:5432/fireData";

	 private static String url = "jdbc:postgresql://192.168.6.202:5432/kgis";
	 private static String username = "dataclean";
	 private static String password = "dataclean";

	public static Connection getConnection() {

		try {
			Class.forName(driver);
			conn = DriverManager.getConnection(url, username, password);
			System.out.println("连接数据库成功!!!");
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e.getClass().getName() + ": " + e.getMessage());
		}
		return conn;

	}

	public static void Disconnect() {
		if (conn != null) {
			try {
				conn.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {
		Util.getConnection();
	}


}
