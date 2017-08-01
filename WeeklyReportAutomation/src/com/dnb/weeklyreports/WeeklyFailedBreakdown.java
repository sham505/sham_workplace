package com.dnb.weeklyreports;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.log4j.Logger;

import com.dnb.utils.AppConstants;
import com.dnb.utils.DBConnection;
import com.dnb.utils.SQLStatements;

/**
 * Description: The responsibility of WeeklyFailedBreakdown class is to fetch failed exports
 * for the given week and exports count by week and month. This class calls generateWeeklyFailedBreakdownWorkbook
 * function to generate report by providing fetched records.
 * 
 * @author shamsheerc
 * @version 1.0
 */
public class WeeklyFailedBreakdown {

	/**
	 * @param args
	 */
	
	public static final Logger log = Logger.getLogger(WeeklyFailedBreakdown.class);
	
	public static void failedBreakdownReport() {
		
		Connection conn = null;
		ResultSet rs = null;	 
		
		try {
			
			log.info("Generation of Weekly Reports started");
			 conn = DBConnection.getConnection(AppConstants.DB_SMSXML);
			 
			 DateFormat dateFormat = new SimpleDateFormat("dd-MMMM-yyyy");
             Date myDate = new Date(System.currentTimeMillis());
             Calendar cal = Calendar.getInstance();
            
            
             cal.setTime(myDate);
             cal.add(Calendar.DATE, -7);
            
            
             int temp = cal.get(Calendar.YEAR);
             int monthNum = cal.get(Calendar.MONTH);
          
            String month = getMonth(monthNum);
            String year = String.valueOf(temp);
            
            System.out.println("Monday Date: "+dateFormat.format(cal.getTime()));
            log.info("Monday Date: "+dateFormat.format(cal.getTime()));
            String date = dateFormat.format(cal.getTime());

		    List<BreakdownVO> breakDownList = new ArrayList<BreakdownVO>();
		    List<BreakdownVO> breakDownWeekList = new ArrayList<BreakdownVO>();
		    List<BreakdownVO> breakDownMonthList = new ArrayList<BreakdownVO>();
		    BreakdownVO vo;
		    log.info("Weekly Failed Reports Query: "+SQLStatements.WeeklyFailedBreakdown(date));
		    PreparedStatement statement = conn.prepareStatement(SQLStatements.WeeklyFailedBreakdown(date));
			                         rs = statement.executeQuery();
			                         
						                      while(rs.next()){
			                        	       vo = new BreakdownVO();
			                        			 vo.setUniqID(rs.getString(1));
			                        			 vo.setDate_created(rs.getDate(2));
			                        			 vo.setApp_name(rs.getString(3));
			                        			 vo.setCustomer_id(rs.getString(4));
			                        			 vo.setEmail(rs.getString(5));
			                        			 vo.setCredits(rs.getString(6));
			                        			 vo.setRows_ordered(rs.getString(7));
			                        			 vo.setProduct_type(rs.getString(8));
			                        			 vo.setExport_format(rs.getString(9));
			                        			 vo.setCustref(rs.getString(10));
			                        			 vo.setWebserver(rs.getString(11));
			                        			 vo.setRetried((rs.getString(12)));
			                        			 breakDownList.add(vo);		 
			                         }
			 log.info("Weekly Breakdown by week query: "+SQLStatements.FailedBreakdownByWeek(date));
			 statement = conn.prepareStatement(SQLStatements.FailedBreakdownByWeek(date));	
			        rs = statement.executeQuery();
			        while(rs.next()){
			        	vo = new BreakdownVO();
			        	vo.setYear(rs.getString(1));
			        	vo.setMonth(rs.getString(2));
			        	vo.setMonth_no(rs.getString(3));
			        	vo.setWeek_no(rs.getString(4));
			        	vo.setApp_name(rs.getString(5));
			        	vo.setTotal_exports(rs.getString(6));
			        	vo.setFailed_exports(rs.getString(7));
			        	breakDownWeekList.add(vo);
			        }  
			        log.info("Weekly Breakdown by month query: "+SQLStatements.FailedBreakdownByMonth(month,year));       
			 statement = conn.prepareStatement(SQLStatements.FailedBreakdownByMonth(month,year));
			        rs = statement.executeQuery();
			        
			        while(rs.next()){
			        	vo = new BreakdownVO();
			        	vo.setYear(rs.getString(1));
			        	vo.setMonth(rs.getString(2));
			        	vo.setMonth_no(rs.getString(3));
			        	vo.setApp_name(rs.getString(4));
			        	vo.setTotal_exports(rs.getString(5));
			        	vo.setFailed_exports(rs.getString(6));
			        	breakDownMonthList.add(vo);
			        	
			        }
			
			        GenerateWorkbook.generateWeeklyFailedBreakdownWorkbook(breakDownList,breakDownWeekList,breakDownMonthList);			        
			
	}catch (Exception e) {
		e.printStackTrace();
		log.error("Exception in fetching weekly breakdown report values ", e);
	}
	finally{
		try {
			rs.close();
			conn.close();
			System.out.println("Connection closed");
			log.info("Connection closed");
		} catch (SQLException e) {
			e.printStackTrace();
			log.error("Exception in fetching weekly breakdown report values ", e);
		}
   	}
}

	private static String getMonth(int monthNum) {
		
		 String month[] = {"January","February","March","April","May","June","July","August","September","October","November","December"};
		
		return month[monthNum];
	}
}
