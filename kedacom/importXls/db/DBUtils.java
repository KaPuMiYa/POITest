package com.kedacom.importXls.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DBUtils {
	
	
	private static String userName ="dataclean";
	
	private static String password ="dataclean";
	
	private static String url="jdbc:postgresql://localhost:5432/test";
	
	public static Connection connection =null;
	
	public static Connection connect(){
		 try {
			Class.forName("org.postgresql.Driver");
			 connection= DriverManager.getConnection(url, userName, password);

			 System.out.println("成功连接pg数据库"+connection);
		} catch (Exception e) {
			
			e.printStackTrace();
		}
         return connection;
        
	}
	public static void Disconnect() {
		if(connection!= null){
			try {
				connection.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {
		DBUtils.connect();
	}


	

}
