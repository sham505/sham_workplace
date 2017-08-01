package com.dnb.weeklyreports;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.TreeMap;

import com.dnb.utils.AppConstants;
import com.dnb.utils.DBConnection;
import com.dnb.utils.SQLStatements;

public class SMSXMLGOLFUsageReport {

	/**
	 * @param args
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void smsxmlGolfReport() {
		
		Connection con = null;
 
		try {
			 con = DBConnection.getConnection(AppConstants.DB_GOLF);
		    DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            Date myDate = new Date(System.currentTimeMillis());
            Calendar cal = Calendar.getInstance();
            int temp = cal.get(Calendar.YEAR);
            String year = String.valueOf(temp);
            cal.setTime(myDate);
            cal.add(Calendar.DATE, -9);
            Calendar cal1 = Calendar.getInstance();
            cal1.setTime(myDate);
            cal1.add(Calendar.DATE, -2);
            System.out.println("Start Date: "+dateFormat.format(cal.getTime()));
            System.out.println("End Date: "+dateFormat.format(cal1.getTime()));
            String startDate = dateFormat.format(cal.getTime());
		    String endDate = dateFormat.format(cal1.getTime());
		    
		    
		   // Map totalUsageMap = new TreeMap();
			Map weeklyMap = new TreeMap();
		    PreparedStatement statement = con.prepareStatement(SQLStatements.GOLFUsageQuery(startDate, endDate));
			System.out.println("Generated Query for SMSXMLGOLF: "+SQLStatements.GOLFUsageQuery(startDate, endDate));
		    ResultSet rs = statement.executeQuery();
			while(rs.next()){
				weeklyMap.put(rs.getString("request_type"),rs.getString("Requests") );
			}
			
			weeklyMap = GenerateWorkbook.formatGOLFUsageMap(weeklyMap);
			
		/*	totalUsageMap = GenerateWorkbook.prepareGOLFTotalUsageMap(weeklyMap,startDate,year);
			
			GenerateWorkbook.generateSMSXMLGOLFWorkBook(totalUsageMap,startDate,endDate);
		*/	
			GenerateWorkbook.generateSMSXMLGOLFWorkBook(weeklyMap);
			
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		finally{
			if(con != null)
				try {
					con.close();
					System.out.println("Connection closed");
				} catch (SQLException e) {
					e.printStackTrace();
				}
		}

	}

}
