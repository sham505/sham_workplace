package com.dnb.utils;

public class SQLStatements {
	
	public static String GRSUsageQuery(String startDate,String endDate){
		return " use SMS_XML_USER " +
				"select (CASE WHEN request_type='AAGGREGATERQ' then 'GRID EXPORT'" +
				" WHEN request_type='AGGREGATERQ' then 'GRID VIEW' " +
				"WHEN request_type='ALOOKUPRQ' then 'EXPORT' " +
				"WHEN request_type='ATREELOOKUPRQ' then 'TREE EXPORT' " +
				"WHEN request_type='DBAICTRQ' then 'DBAI CLICK THROUGH' " +
				"WHEN request_type='INDEXRQ' then 'COUNT SEARCH' " +
				"WHEN request_type='LOOKUPRQ' then 'REPORT VIEW' " +
				"WHEN request_type='PRINT-DMLIST' then 'DIRECT MARKETING LIST PRINT' " +
				"WHEN request_type='PRINT-LIST' then 'LIST PRINT' " +
				"WHEN request_type='PRINT-RPT' then 'REPORT PRINT' " +
				"WHEN request_type='PRINT-TREE' then 'TREE PRINT' " +
				"WHEN request_type='PRINT-TREEPLUS' then 'TREEPLUS PRINT' " +
				"WHEN request_type='TREELOOKUPRQ' then 'TREE VIEW' " +
				"WHEN request_type='SMSLOGINRQ' then 'TOTAL LOGIN' " +
				"ELSE request_type END) " +
				"request_type, count(*) as Requests from usage_7days with (NOLOCK) " +
				"where convert(varchar,request_date,121) between '"+startDate+ " 00:00:00.000' and '"+endDate+"  23:59:59.999' " +
				"group by request_type order by request_type";
	}
	
	public static String GOLFUsageQuery(String startDate,String endDate){
		return " use SMS_XML_USER " +
		"select (CASE WHEN request_type='AAGGREGATERQ' then 'GRID EXPORT'" +
		" WHEN request_type='AGGREGATERQ' then 'GRID VIEW' " +
		"WHEN request_type='ALOOKUPRQ' then 'EXPORT' " +
		"WHEN request_type='ATREELOOKUPRQ' then 'TREE EXPORT' " +
		"WHEN request_type='DBAICTRQ' then 'DBAI CLICK THROUGH' " +
		"WHEN request_type='INDEXRQ' then 'COUNT SEARCH' " +
		"WHEN request_type='LOOKUPRQ' then 'REPORT VIEW' " +
		"WHEN request_type='PRINT-DMLIST' then 'DIRECT MARKETING LIST PRINT' " +
		"WHEN request_type='PRINT-LIST' then 'LIST PRINT' " +
		"WHEN request_type='PRINT-RPT' then 'REPORT PRINT' " +
		"WHEN request_type='PRINT-TREE' then 'TREE PRINT' " +
		"WHEN request_type='PRINT-TREEPLUS' then 'TREEPLUS PRINT' " +
		"WHEN request_type='TREELOOKUPRQ' then 'TREE VIEW' " +
		"WHEN request_type='SMSLOGINRQ' then 'TOTAL LOGIN' " +
		"ELSE request_type END) " +
		"request_type, count(*) as Requests from usage_7days with (NOLOCK) " +
		"where convert(varchar,request_date,121) between '"+startDate+ " 00:00:00.000' and '"+endDate+"  23:59:59.999' " +
		"group by request_type order by request_type";
	}
	
	public static String WeeklyFailedBreakdown(String date){
		return "use sms_xml_user " +
				"select o.uniqid, o.date_created, o.app_name, o.customer_id, o.email, " +
				"o.credits, o.rows_ordered,o.product_type, o.export_format, o.custref, " +
				"'web' as webserver, 'N' as retried " +
				"from sms_xml_user.dbo.orderrq o with(NOLOCK)," +
				" sms_xml_user.dbo.xmlinterfacev2_xmlrq x with(NOLOCK) " +
				"where o.date_created > '"+date+"' and o.complete <> 1  " +
						"and o.email <> 'ewowcanada@dnb.com' and o.email not like '%@csc.com' " +
						"and x.ref = o.uniqid";
	}
	
	public static String FailedBreakdownByWeek(String date){
		return "use sms_xml_user " +
				"select	datename(year,o1.date_created) as year," +
				" datename(month,o1.date_created) as month," +
				"datepart(month,o1.date_created) as month_no," +
				"datename(week,o1.date_created) as week_no," +
				"o1.app_name,count(o1.complete) as total_exports," +
				"count(o2.complete) as failed_exports from " +
				"sms_xml_user.dbo.orderrq o1 with (NOLOCK) " +
				"left outer join (select * from sms_xml_user.dbo.orderrq with (NOLOCK)" +
				" where complete <> 1 and email <> 'ewowcanada@dnb.com' and email" +
				"  not like '%@csc.com') o2 on o1.uniqid= o2.uniqid " +
				"where o1.date_created > '"+date+"' " +
						"group by  datename(year,o1.date_created), " +
						"datepart(month,o1.date_created), " +
						"datename(week,o1.date_created), " +
						"datename(month,o1.date_created), " +
						"o1.app_name order by  datename(year,o1.date_created), " +
						"datepart(month,o1.date_created), " +
						"datename(week,o1.date_created), " +
						"datename(month,o1.date_created), o1.app_name";
	}
	
	public static String FailedBreakdownByMonth(String month,String year){
		return "use sms_xml_user " +
				"select	datename(year,o1.date_created) as year," +
				" datename(month,o1.date_created) as month," +
				"datepart(month,o1.date_created) as month_no," +
				"o1.app_name,count(o1.complete) as total_exports," +
				"count(o2.complete) as failed_exports from sms_xml_user.dbo.orderrq o1 with (NOLOCK)" +
				"left outer join (select * from sms_xml_user.dbo.orderrq with (NOLOCK) " +
				"where complete <> 1 and email <> 'ewowcanada@dnb.com' and email  not like '%@csc.com') " +
				"o2 on o1.uniqid= o2.uniqid " +
				"where o1.date_created > '01 "+month+" "+year+"' " +
						"group by  datename(year,o1.date_created), " +
						"datepart(month,o1.date_created), " +
						"datename(month,o1.date_created), " +
						"o1.app_name order by  datename(year,o1.date_created), " +
						"datepart(month,o1.date_created), " +
						"datename(month,o1.date_created), o1.app_name";
	}
	

}
