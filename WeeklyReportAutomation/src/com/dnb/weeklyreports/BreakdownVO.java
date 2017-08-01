package com.dnb.weeklyreports;

import java.util.Date;

public class BreakdownVO {
	
	public String uniqID;
	public Date date_created;
	public String app_name;
	public String customer_id;
	public String email;
	public String credits;
	public String rows_ordered;
	public String product_type;
	public String export_format;
	public String custref;
	public String webserver;
	public String retried;
	public String year;
	public String month;
	public String month_no;
	public String week_no;
	public String total_exports;
	public String failed_exports;
	
	
	public BreakdownVO(){
		
	}
	
	public BreakdownVO(String uniqID,Date date_created, String app_name, String customer_id, String email, String credits, String rows_ordered, String product_type, String export_format, String custref, String webserver, String retried){
		this.uniqID = uniqID;
		this.date_created = date_created;
		this.app_name = app_name;
		this.customer_id = customer_id;
		this.email = email;
		this.credits = credits;
		this.rows_ordered = rows_ordered;
		this.product_type = product_type;
		this.export_format = export_format;
		this.custref = custref;
		this.webserver = webserver;
		this.retried = retried;
	}
	
	
	public String getUniqID() {
		return uniqID;
	}
	public void setUniqID(String uniqID) {
		this.uniqID = uniqID;
	}
	public Date getDate_created() {
		return date_created;
	}
	public void setDate_created(Date date_created){
		
		this.date_created = date_created;
	}
	public String getApp_name() {
		return app_name;
	}
	public void setApp_name(String app_name) {
		this.app_name = app_name;
	}
	public String getCustomer_id() {
		return customer_id;
	}
	public void setCustomer_id(String customer_id) {
		this.customer_id = customer_id;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getCredits() {
		return credits;
	}
	public void setCredits(String credits) {
		this.credits = credits;
	}
	public String getRows_ordered() {
		return rows_ordered;
	}
	public void setRows_ordered(String rows_ordered) {
		this.rows_ordered = rows_ordered;
	}
	public String getProduct_type() {
		return product_type;
	}
	public void setProduct_type(String product_type) {
		this.product_type = product_type;
	}
	public String getExport_format() {
		return export_format;
	}
	public void setExport_format(String export_format) {
		this.export_format = export_format;
	}
	public String getCustref() {
		return custref;
	}
	public void setCustref(String custref) {
		this.custref = custref;
	}
	public String getWebserver() {
		return webserver;
	}
	public void setWebserver(String webserver) {
		this.webserver = webserver;
	}
	public String getRetried() {
		return retried;
	}
	public void setRetried(String retried) {
		this.retried = retried;
	}
	public String getYear() {
		return year;
	}

	public void setYear(String year) {
		this.year = year;
	}

	public String getMonth() {
		return month;
	}

	public void setMonth(String month) {
		this.month = month;
	}

	public String getMonth_no() {
		return month_no;
	}

	public void setMonth_no(String month_no) {
		this.month_no = month_no;
	}

	public String getWeek_no() {
		return week_no;
	}

	public void setWeek_no(String week_no) {
		this.week_no = week_no;
	}

	public String getTotal_exports() {
		return total_exports;
	}

	public void setTotal_exports(String total_exports) {
		this.total_exports = total_exports;
	}

	public String getFailed_exports() {
		return failed_exports;
	}

	public void setFailed_exports(String failed_exports) {
		this.failed_exports = failed_exports;
	}

	
}
