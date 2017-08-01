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

import org.apache.log4j.Logger;

import com.dnb.utils.AppConstants;
import com.dnb.utils.DBConnection;
import com.dnb.utils.SQLStatements;

public class SMSXMLUsageReport {

	/**
	 * @param args
	 */
	public static final Logger log = Logger.getLogger(SMSXMLUsageReport.class);
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void smsxmlUsageReport() {
		
		Connection con = DBConnection.getConnection(AppConstants.DB_SMSXML);
		
		try {
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
            
            String startDate = dateFormat.format(cal.getTime());
		    String endDate = dateFormat.format(cal1.getTime());
		   
			Map weeklyMap = new TreeMap();
		    PreparedStatement statement = con.prepareStatement(SQLStatements.GRSUsageQuery(startDate, endDate));
		    log.info("generated query for SMSXML Usage report: "+SQLStatements.GRSUsageQuery(startDate, endDate));
		    System.out.println("Generated Query for GRS Usage: "+(SQLStatements.GRSUsageQuery(startDate, endDate)));
			ResultSet rs = statement.executeQuery();
			while(rs.next()){
				weeklyMap.put(rs.getString("request_type"),rs.getString("Requests") );
			}
			weeklyMap     = GenerateWorkbook.formatGRSUsageMap(weeklyMap);
			
		//	yearlyUsageMap = GenerateWorkbook.prepareGRSTotalUsageMap(weeklyMap,startDate,year);
			
			
			/**
			 * Calling Function to generate SMSXML usage report 
			 */
	//		GenerateWorkbook.generateSMSXMLWorkBook(yearlyUsageMap,startDate,endDate);
			
				
			
			GenerateWorkbook.generateSMSXMLWorkBook(weeklyMap);
			
			
			/*Set s = usageMap.entrySet();
			Iterator it = s.iterator();
			
			while(it.hasNext()){
				Map.Entry me = (Map.Entry)it.next();
			System.out.print("key: "+me.getKey());
			System.out.println("value: "+me.getValue());
			}*/
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
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
