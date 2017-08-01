package com.dnb.utils;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import org.apache.log4j.Logger;


public class DBConnection {
	
	public static final Logger log = Logger.getLogger(DBConnection.class);
	
	 
	public static Connection getConnection(String dbserver){
		
		Connection connection = null;
		String connectionUrl = "";
        try
        {
            Class.forName(AppConstants.DB_DRIVER);
            if(dbserver.equals("PPE"))
            connectionUrl = getConnUrls("PPE_ConnectionURL");
            else if(dbserver.equals("SMSXML"))
                connectionUrl = getConnUrls("SMSXML_ConnectionURL");
            else if(dbserver.equals("SMSXMLGOLF"))
                connectionUrl = getConnUrls("SMSXMLGOLF_ConnectionURL");   
            String userId = AppConstants.userID;
            String password = AppConstants.Password;
            connection = DriverManager.getConnection(connectionUrl, userId, password);
            if(connection !=null){
            	System.out.println("connection is established");
            }
        }
        catch(SQLException e)
        {
        	log.error("Exception in establishing SQL connection", e);
            System.out.println((new StringBuilder("Sql Exception")).append(e).toString());
        }
        catch(Exception e)
        {
        	log.error("Exception in establishing connection", e);
            System.out.println((new StringBuilder()).append(e).toString());
        }
        return connection;
	}
	
	public static String getConnUrls(String URL){
		if(URL.equals("SMSXML_ConnectionURL"))
			return AppConstants.SMSXML_ConnectionURL;
		if(URL.equals("SMSXMLGOLF_ConnectionURL"))
			return AppConstants.SMSXMLGOLF_ConnectionURL;
		if(URL.equals("PPE_ConnectionURL"))
			return AppConstants.PPE_ConnectionURL;
		else return URL;
	}
	
	

}
