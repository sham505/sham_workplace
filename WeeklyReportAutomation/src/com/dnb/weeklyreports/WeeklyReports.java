package com.dnb.weeklyreports;

import org.apache.log4j.Logger;




public class WeeklyReports {

	/**
	 * @param args
	 * 
	 */
	
	final static Logger log = Logger.getLogger(WeeklyReports.class);
	
	public static void main(String[] args) {
		
	//	BasicConfigurator.configure();
		
		

		WeeklyFailedBreakdown.failedBreakdownReport();
		SMSXMLUsageReport.smsxmlUsageReport();
		SMSXMLGOLFUsageReport.smsxmlGolfReport();
		
		
		
		
		

	}

}
