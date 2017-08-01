package com.dnb.weeklyreports;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Description: GenerateWorkBook class generates all the weekly reports by iterating through the 
 * data available from respective report classes (SMSXMLUsageReport.java, SMSXMLGOLFUsageReport.java,
 * WeeklyFailedBreakdownReport.java). This class also takes the responsibility of sending reports via 
 * to the provided distribution list.
 * 
 * @author shamsheer
 * @version 1.0
 * @implements Serializable
 */

public class GenerateWorkbook implements Serializable{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	public static final Logger log = Logger.getLogger(GenerateWorkbook.class);

	public static String sms_xml_request_type[] = { "AMINLOOKUPRQ",
			"APLUSLOOKUPRQ", "ATREELOOKUPRQ", "ATREEPLUSLOOKUPRQ",
			"AVIABLOOKUPRQ","CHGALRTRPTRQ", "COUNT SEARCH", "DBAI CLICK THROUGH",
			"DIRECT MARKETING LIST PRINT", "EXPORT", "EXPORT-MINSRPT",
			"EXPORT-MINTREERPT", "EXPORT-OWNERSRPT", "GRID EXPORT",
			"GRID VIEW", "GRSMINRPT", "LIST PRINT", "MINSRPT", "MINTREERPT",
			"OWNERSRPT", "PRINT-CHALRTLIST", "PRINT-GRSMINRPT",
			"PRINT-VIAB-RPT", "PRINT-MINSRPT", "PRINT-MINTREERPT",
			"PRINT-OWNERSRPT", "REPORT PRINT", "REPORT VIEW", "TOTAL LOGIN",
			"TREE EXPORT", "TREE PRINT", "TREE VIEW", "TREEPLUS PRINT",
			"TREEPLUSLOOKUPRQ"};
	
	public static String sms_xml_golf_request_type[] = {"COUNT SEARCH", "GBOLOOKUPRQ",
			"REPORT VIEW", "TREE VIEW", "TREEPLUSLOOKUPRQ" };

	/**
	 * This method is called generates SMSXML weekly report.
	 * @param weeklyMap
	 */
	
	@SuppressWarnings("rawtypes")
	public static void generateSMSXMLWorkBook(Map weeklyMap){
		
		DateFormat dateFormat = new SimpleDateFormat("ddMMMyyyy");
		Date myDate = new Date(System.currentTimeMillis());
		Calendar cal = Calendar.getInstance();
		final String year = String.valueOf(cal.get(Calendar.YEAR));
		cal.setTime(myDate);
        cal.add(Calendar.DATE, -3);
        
        String lastReportDate = getDate(10,"ddMMMyyyy");
        
        String lastGeneratedreport = getDate(10,"dd-MMM");
        
        String time = dateFormat.format(cal.getTime());
        
        HSSFSheet sheet;
        HSSFWorkbook workbook = null;
        
        HSSFRow row;
        HSSFRow currentRow;
        HSSFCell cell;
     
		try{
        FileInputStream input = new FileInputStream(new File("Weekly_GRS_usage_Week_Ending_"+lastReportDate+".xls"));
		             workbook = new HSSFWorkbook(input);
        
        sheet = workbook.getSheet(year);      
        
        int currentColumn  = 0;
		int CurrentRowNo = 0;
		int noOfCells = 0;
        
        if(sheet == null){
        	log.info("SMSXMLUsage sheet unavailable, crating new sheet");
        	Set set = weeklyMap.entrySet();
    		Iterator it = set.iterator();

    		sheet = workbook.createSheet(year);
    		
        	row   = sheet.createRow(0);
        	cell  = row.createCell(0);
        	cell.setCellValue("Week Ending");
        	cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

        	cell = row.createCell(1);
			cell.setCellValue("Request_Type");
			cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

			cell = row.createCell(2);
			cell.setCellValue(getDate(3, "dd-MMM"));
			cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

			CurrentRowNo++;		 
			currentColumn = 1;
			int firstColumn = 0;

			while(it.hasNext()){
				 
				 Map.Entry weeklyEntry = (Map.Entry) it.next();
				 row = sheet.createRow(CurrentRowNo);
				
				 if(firstColumn == 17){
					 cell = row.createCell(currentColumn-1);
					 cell.setCellValue("Request_Type");
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 }
				 else{
					 cell = row.createCell(currentColumn-1);
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 }
				 cell = row.createCell(currentColumn);
				 cell.setCellValue((String)weeklyEntry.getKey());
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 cell = row.createCell(currentColumn+1, HSSFCell.CELL_TYPE_NUMERIC);
				 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 CurrentRowNo++;
				 firstColumn++;
			 }
			log.info("Iterated throgh smsxml usage values and written to sheet");
			for (int j=0; j<12; j++){
				   sheet.autoSizeColumn(j);
				}
			workbook.setSheetOrder(year, 0);
            FileOutputStream out = new FileOutputStream("Weekly_GRS_usage_Week_Ending_"+time+".xls");
			workbook.write(out);
			out.close();	
        log.info("SMSXMLUsage report generated successfully");
        }
        else{
        log.info("Iterating through the sheet values to find the last report usage details");
        int noOfRows = 	sheet.getLastRowNum();
				
		Set set = weeklyMap.entrySet();
        
		Iterator it = set.iterator();
		
		if(CurrentRowNo == 0){	
        
        for (int rowNum =0; rowNum < noOfRows; rowNum++){
			
    	currentRow = sheet.getRow(rowNum);	
    	
    	if(currentRow != null)
    	{
    	  noOfCells = currentRow.getLastCellNum();
			
			 for(int colNum =0; colNum < noOfCells; colNum ++){
				
				cell = sheet.getRow(rowNum).getCell(colNum);
				
			    	if(cell != null){
					
				         if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
					
					         if(cell.getStringCellValue().equalsIgnoreCase(lastGeneratedreport)){
						        		 CurrentRowNo = rowNum;
							             break;
					  }
				   }
			    }
			 }
          }
			if(CurrentRowNo != 0){
				break;
		   }	
		  }       
		 }      
		 currentColumn = sheet.getRow(CurrentRowNo).getLastCellNum(); 
		log.info("Iterated though the sheet and found the column no: "+currentColumn);
		 if(currentColumn > 13){ 
			 log.info("As current column is more than 13 appending the report to the next row");
			 CurrentRowNo = noOfRows + 2;
			 int firstColumn = 0;
			 row  = sheet.createRow(CurrentRowNo);			 
			 cell = row.createCell(0);
			 cell.setCellValue("Week Ending");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(1);
			 cell.setCellValue("Request_Type");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(2);
			 cell.setCellValue(getDate(3, "dd-MMM"));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 CurrentRowNo++;		 
			 currentColumn = 1;
			 
			 while(it.hasNext()){
				 
				 Map.Entry weeklyEntry = (Map.Entry) it.next();
				 
				 row = sheet.createRow(CurrentRowNo);
				
				 if(firstColumn == 17){
					 cell = row.createCell(currentColumn-1);
					 cell.setCellValue("Request_Type");
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 }
				 else{
					 cell = row.createCell(currentColumn-1);
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 }
				 cell = row.createCell(currentColumn);
				 cell.setCellValue((String)weeklyEntry.getKey());
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 cell = row.createCell(currentColumn+1, HSSFCell.CELL_TYPE_NUMERIC);
				 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 CurrentRowNo++;
				 firstColumn++;
			 }
			 log.info("Iterated and appended values to the sheet");
			 FileOutputStream out = new FileOutputStream("Weekly_GRS_usage_Week_Ending_"+time+".xls");
				workbook.write(out);
				out.close();		 
			 log.info("SMSXMLUsage report genetared successfully");
		 }else{
			 row  = sheet.getRow(CurrentRowNo);
			 cell = row.createCell(currentColumn);
			 cell.setCellValue(getDate(3, "dd-MMM"));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 CurrentRowNo = CurrentRowNo + 1;

			 while(it.hasNext()){
					 Map.Entry weeklyEntry = (Map.Entry) it.next();
					 row  = sheet.getRow(CurrentRowNo);
					 cell = row.createCell(currentColumn, HSSFCell.CELL_TYPE_NUMERIC);
					 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
					 CurrentRowNo ++;
			 }
			 log.info("Iterated and appended values to the sheet");
			 FileOutputStream out = new FileOutputStream("Weekly_GRS_usage_Week_Ending_"+time+".xls");
				workbook.write(out);
				out.close();
				log.info("SMSXMLUsage report genetared successfully");
		 }
        }
		 String mailBody = "<FONT style = 'font-family:Calibri'>Hello All,<BR><BR>Please find the Weekly GRS Usage Report for Week ending "+time+"</FONT>";
	     String subject = "Weekly GRS Usage Report for Week ending "+time;
	     String filePath = "Weekly_GRS_usage_Week_Ending_"+time+".xls";
		 String toAddress = "shamsheerc@dnb.com";
	     GenerateMail.sendMail(mailBody, subject, filePath, toAddress);
	     log.info("SMSXML Usage report generated and sent successfully");
		}catch (IOException e) {
			log.error("Exception in generating SMSXML usage report", e);
			e.printStackTrace();
		}
	}
	
	@SuppressWarnings("rawtypes")
	public static void generateSMSXMLGOLFWorkBook(Map weeklyMap){
		
		DateFormat dateFormat = new SimpleDateFormat("ddMMMyyyy");
		Date myDate = new Date(System.currentTimeMillis());
		Calendar cal = Calendar.getInstance();
		final String year = String.valueOf(cal.get(Calendar.YEAR));
		cal.setTime(myDate);
        cal.add(Calendar.DATE, -3);
        
        String lastReportDate = getDate(10,"ddMMMyyyy");
        
        String lastGeneratedreport = getDate(10,"dd-MMM");
        
        String time = dateFormat.format(cal.getTime());
        
        HSSFSheet sheet;
        HSSFWorkbook workbook = null;
        
        HSSFRow row;
        HSSFRow currentRow;
        HSSFCell cell;

        try{
        FileInputStream input = new FileInputStream(new File("Weekly_SMSXMLGOLF_usage_Week_Ending_"+lastReportDate+".xls"));
		             workbook = new HSSFWorkbook(input);
        
        sheet = workbook.getSheet(year);      
        int currentColumn  = 0;
		int CurrentRowNo = 0;
		int noOfCells = 0;
		
		if(sheet == null){
			log.info("SMSXMLGOLFUsage sheet unavailable, crating new sheet");
			 sheet = workbook.createSheet(year);
			 row  = sheet.createRow(CurrentRowNo);			 
			 cell = row.createCell(0);
			 cell.setCellValue("Week");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

			 cell = row.createCell(1);
			 cell.setCellValue(getDate(3, "dd-MMM"));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

			 CurrentRowNo++;		 
			 currentColumn = 0;
			
			 Set set = weeklyMap.entrySet();
			 Iterator it = set.iterator();
				
			 while(it.hasNext()){
				 
				 Map.Entry weeklyEntry = (Map.Entry) it.next();
				 row = sheet.createRow(CurrentRowNo);
				 cell = row.createCell(currentColumn);
				 cell.setCellValue((String)weeklyEntry.getKey());
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 cell = row.createCell(currentColumn+1, HSSFCell.CELL_TYPE_NUMERIC);
				 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 CurrentRowNo++;	
			 }
			 log.info("Iterated throgh smsxmlgolf usage values and written to sheet");
			 for (int j=0; j<12; j++){
				   sheet.autoSizeColumn(j);
				}
			 workbook.setSheetOrder(year, 0);
			 FileOutputStream out = new FileOutputStream("Weekly_SMSXMLGOLF_usage_Week_Ending_"+time+".xls");
				workbook.write(out);
				out.close();
				log.info("SMSXMLUsage report genetared successfully");
		}
		else{
        log.info("Iterating to find the last report usage details");
		int noOfRows = 	sheet.getLastRowNum();
				
		Set set = weeklyMap.entrySet();
        
		Iterator it = set.iterator();
		
	    if(CurrentRowNo == 0){	
        
        for (int rowNum =0; rowNum < noOfRows; rowNum++){
			
    	currentRow = sheet.getRow(rowNum);	
    	
    	if(currentRow != null)
    	{
    	  noOfCells = currentRow.getLastCellNum();
			
			 for(int colNum =0; colNum < noOfCells; colNum ++){
				
				cell = sheet.getRow(rowNum).getCell(colNum);
				
			    	if(cell != null){
					
				         if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
					
					         if(cell.getStringCellValue().equalsIgnoreCase(lastGeneratedreport)){
						        		 CurrentRowNo = rowNum;
							             break;
					  }
				   }
			    }
			 }
          }
			if(CurrentRowNo != 0){
				break;
		   }	
		  }       
		 }      
		 currentColumn = sheet.getRow(CurrentRowNo).getLastCellNum(); 
		log.info("found the last week's usage details column:  "+currentColumn);
		 if(currentColumn > 12){ 
			 log.info("As current column is more than 12 writing the values to the next row");
			 CurrentRowNo = noOfRows + 2;		 
			 row  = sheet.createRow(CurrentRowNo);			 
			 cell = row.createCell(0);
			 cell.setCellValue("Week");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(1);
			 cell.setCellValue(getDate(3, "dd-MMM"));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 CurrentRowNo++;		 
			 currentColumn = 0;
			 
			 while(it.hasNext()){
				 
				 Map.Entry weeklyEntry = (Map.Entry) it.next();
				 row = sheet.createRow(CurrentRowNo);
				 cell = row.createCell(currentColumn);
				 cell.setCellValue((String)weeklyEntry.getKey());
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 cell = row.createCell(currentColumn+1, HSSFCell.CELL_TYPE_NUMERIC);
				 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				 CurrentRowNo++;	
			 }
			 log.info("SMSXMLGOLF usage report values has been appended to the sheet");
			 FileOutputStream out = new FileOutputStream("Weekly_SMSXMLGOLF_usage_Week_Ending_"+time+".xls");
				workbook.write(out);
				out.close();	
				log.info("SMSXMLGOLF usage report generated successfully");
			 
		 }else{
			 row  = sheet.getRow(CurrentRowNo);
			 cell = row.createCell(currentColumn);
			 cell.setCellValue(getDate(3, "dd-MMM"));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 CurrentRowNo = CurrentRowNo + 1;

			 while(it.hasNext()){
				 
					 Map.Entry weeklyEntry = (Map.Entry) it.next();
					 row  = sheet.getRow(CurrentRowNo);
					 cell = row.createCell(currentColumn, HSSFCell.CELL_TYPE_NUMERIC);
					 cell.setCellValue(Integer.parseInt((String)weeklyEntry.getValue()));
					 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
					 CurrentRowNo ++;
		        	 
			 }
			 log.info("SMSXMLGOLF usage report values has been appended to the sheet");
			 FileOutputStream out = new FileOutputStream("Weekly_SMSXMLGOLF_usage_Week_Ending_"+time+".xls");
				workbook.write(out);
				out.close();
				log.info("SMSXMLGOLF usage report generated successfully");
		 }
		}
		 
		String mailBody = "<FONT style = 'font-family:Calibri'>Hello All,<BR><BR>Please find the Weekly SMSXMLGOLF Usage Report for Week ending "+time+"</FONT>";
	    String subject = "Weekly SMSXMLGOLF Usage Report for Week ending  "+time;
	    String filePath = "Weekly_SMSXMLGOLF_usage_Week_Ending_"+time+".xls";
	    String toAddress = "shamsheerc@dnb.com";

	     GenerateMail.sendMail(mailBody, subject, filePath, toAddress);
	     log.info("SMSXMLGOLF Usage report generated and sent successfully");
		}catch (IOException e) {
			log.error("Exception in generating SMSXMLGOLF usage report", e);
			throw new RuntimeException();
		}
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static Map formatGRSUsageMap(Map weeklyMap) {

		log.info("Formatting smsxml usage map");
		Set set = weeklyMap.entrySet();
		Iterator iterator = set.iterator();
		String zero = "0";

		ArrayList list = new ArrayList(Arrays.asList(sms_xml_request_type));

		while (iterator.hasNext()) {
			Map.Entry entry = (Map.Entry) iterator.next();
			String key = (String) entry.getKey();

			if (list.contains(key)) {
				list.remove(key);
			}
		}
		Iterator it = list.iterator();
		while (it.hasNext()) {
			weeklyMap.put((String) it.next(), zero);
		}
		if(weeklyMap.containsKey("GBOLOOKUPRQ")){
			weeklyMap.remove("GBOLOOKUPRQ");
		}
		return weeklyMap;
	}
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static Map formatGOLFUsageMap(Map weeklyMap) {
		log.info("Formatting smsxmlgolf usage map");
		Set set = weeklyMap.entrySet();
		Iterator iterator = set.iterator();
		String zero = "0";
		ArrayList list = new ArrayList(Arrays.asList(sms_xml_golf_request_type));

		while (iterator.hasNext()) {
			Map.Entry entry = (Map.Entry) iterator.next();
			String key = (String) entry.getKey();

			if (list.contains(key)) {
				list.remove(key);
			}
		}
		Iterator it = list.iterator();
		while (it.hasNext()) {
			weeklyMap.put((String) it.next(), zero);
		}
		return weeklyMap;
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static HSSFCellStyle formatWorkBookCells(String type,HSSFWorkbook workbook ){
		
		HSSFCellStyle Headerstyle = workbook.createCellStyle();
		Headerstyle.setFillForegroundColor(HSSFColor.YELLOW.index);
		Headerstyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Headerstyle.setBorderBottom((short)1);
		Headerstyle.setBorderTop((short)1);
		Headerstyle.setBorderRight((short)1);
		Headerstyle.setBorderLeft((short)1);
		
		HSSFFont font = workbook.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		Headerstyle.setFont(font);
		HSSFCellStyle cellstyle = workbook.createCellStyle();
		cellstyle.setBorderBottom((short)1);
		cellstyle.setBorderTop((short)1);
		cellstyle.setBorderRight((short)1);
		cellstyle.setBorderLeft((short)1);
		
		
		HashMap m = new HashMap();
		m.put("HeaderStyle", Headerstyle);
		m.put("cellStyle", cellstyle);

		return (HSSFCellStyle)m.get(type);
	}
	
	/**
	 * 
	 * Weekly Failed Breakdown Report:  
	 * This function takes failed report values as list and appends the values to the previous report 
	 * @author Shamsheer
	 * @param breakDownList - Failed reports for last week
	 * @param breakDownWeekList - Failed reports by week
	 * @param breakDownMonthList - Failed reports by month
	 */

	@SuppressWarnings("rawtypes")
	public static void generateWeeklyFailedBreakdownWorkbook(
			List<BreakdownVO> breakDownList,List<BreakdownVO> breakDownWeekList, List<BreakdownVO> breakDownMonthList) {
		
		try{
			 DateFormat dateFormat = new SimpleDateFormat("ddMMMyyyy");
             Date myDate = new Date(System.currentTimeMillis());
             Calendar cal = Calendar.getInstance();
             cal.setTime(myDate);
             String year = String.valueOf(cal.get(Calendar.YEAR));
             cal.add(Calendar.DATE, -14);
             String previousWorkBookDate = dateFormat.format(cal.getTime());
             
		FileInputStream input = new FileInputStream(new File("Weekly_Failed_breakdown_Week_Starting_"+previousWorkBookDate+".xls"));
		HSSFWorkbook workbook = new HSSFWorkbook(input);
		HSSFSheet sheet = workbook.getSheet("Failed this Week");
		
		if(sheet == null){
			
			sheet = workbook.createSheet("Failed this Week");
			workbook.setSheetOrder("Failed this Week", 0);
			log.info("Sheet unavailable. Created new sheet");
		}
		HSSFRow row;
		HSSFCell cell;
		BreakdownVO breakdownVO;
		
		Iterator it = breakDownList.iterator();
		int column  = 0;

		for(int i = 1; i <= sheet.getLastRowNum(); i++){
		sheet.createRow(i);	
		}

		int i =1;	
		while(it.hasNext()){
		breakdownVO = (BreakdownVO)it.next();	
		 row  = sheet.createRow(i);
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getUniqID());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(convertDate(breakdownVO.getDate_created()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getApp_name());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getCustomer_id()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getEmail());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getCredits()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getRows_ordered()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getProduct_type());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getExport_format());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getCustref());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getWebserver());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getRetried());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column = 0;
		 i++;
		}
        log.info("Iterated failed reports and set values to sheet");
		for (int j=0; j<12; j++){
			   sheet.autoSizeColumn(j);
			}
		
		input.close();
		String currentWorkBookDate = getDate(7,"ddMMMyyyy");
		FileOutputStream fos = new FileOutputStream("Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls");	
	    workbook.write(fos);
	    fos.close();
	    log.info("failed exports sheet generated successfully and written to workbook");
	    generateFailedBreakdownByWeekWorkbook(breakDownWeekList,year); 
	    generateFailedBreakdownByMonthWorkbook(breakDownMonthList,year);
	    
		}catch (Exception e) {
			log.error("Exception in generating weeklyFailed Breakdown report", e);
			System.out.println("Exception occured while generating failed exports sheet");
			e.printStackTrace();
		}
	}
	
    /**
     * This method generates exports count by week
     * @param breakDownWeekList
     * @param year
     */
	@SuppressWarnings("rawtypes")
	public static void generateFailedBreakdownByWeekWorkbook(List<BreakdownVO> breakDownWeekList,String year){
		
		try{
		String currentWorkBookDate = getDate(7,"ddMMMyyyy");	
		FileInputStream input = new FileInputStream(new File("Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls"));
		HSSFWorkbook workbook = new HSSFWorkbook(input);
		
		HSSFRow row;
		HSSFCell cell;
		BreakdownVO breakdownVO;
		
		
		HSSFSheet weekSheet = workbook.getSheet(year+" - Week");
		
		if(weekSheet == null){
            log.info("weekly failed breakdown by week sheet is not available, creating new sheet");
			weekSheet = workbook.createSheet(year+" - Week");
			
			 row  = weekSheet.createRow(0);
			 cell = row.createCell(0);
			 cell.setCellValue("Year");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(1);
			 cell.setCellValue("Month");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(2);
			 cell.setCellValue("Month No");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(3);
			 cell.setCellValue("Week No");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(4);
			 cell.setCellValue("App Name");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(5);
			 cell.setCellValue("Total Exports");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 cell = row.createCell(6);
			 cell.setCellValue("Failed Exports");
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			 
			workbook.setSheetOrder(year+" - Week", 1);
		}
		                 
		Iterator it = breakDownWeekList.iterator();
		
		int column  = 0;
		int CurrentRowNo = 0;
		
		while(it.hasNext()){
			breakdownVO = (BreakdownVO)it.next();
		
		int noOfRows = 	weekSheet.getLastRowNum();
		
	  if(CurrentRowNo == 0){
		  
		for (int rowNum =0; rowNum < noOfRows; rowNum++){
			
			int noOfCells = weekSheet.getRow(rowNum).getLastCellNum();
			
			for(int colNum =0; colNum < noOfCells; colNum ++){
				
				cell = weekSheet.getRow(rowNum).getCell(colNum);
				
			    	if(cell != null){
					
				         if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
					
					         if(cell.getNumericCellValue() == Double.parseDouble(breakdownVO.getWeek_no())){
						
						         cell = weekSheet.getRow(rowNum).getCell(colNum+1);
						          
						         if(cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
						      
						        	 if(cell.getStringCellValue().equalsIgnoreCase("GRS") || cell.getStringCellValue().equalsIgnoreCase("MIS")){
						        
						        		 CurrentRowNo = rowNum;
							             break;
						   }
						 }
					  }
				   }
			    }
			 }
			if(CurrentRowNo != 0){
				break;
			}	
		  }
		
		if(CurrentRowNo == 0){
			CurrentRowNo = weekSheet.getLastRowNum()+ 1;
		}
	  }
	  
		 row  = weekSheet.createRow(CurrentRowNo);
		 cell = row.createCell(column,HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getYear()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getMonth());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column,HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getMonth_no()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column,HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getWeek_no()));;
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column);
		 cell.setCellValue(breakdownVO.getApp_name());
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column,HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getTotal_exports()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column++;
		 cell = row.createCell(column,HSSFCell.CELL_TYPE_NUMERIC);
		 cell.setCellValue(Integer.parseInt(breakdownVO.getFailed_exports()));
		 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
		 column = 0;
		 CurrentRowNo++;
			}
		log.info("Iterated through breakdown values by week and appended the values to sheet");
		
		for (int j=0; j<7; j++){
			   weekSheet.autoSizeColumn(j);
			}
		input.close();	
	    
		FileOutputStream fos = new FileOutputStream("Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls");	
	    workbook.write(fos);
	    fos.close();
	    log.info("weeklybreakdown by week sheet generated successfully and written to the workbook ");
	    
	    }catch (Exception e) {
	    	log.error("exception in generating weeklybreakdown by week sheet", e);
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		
		
	}
	/**
	 * This method provides exports count by month
	 * @param breakDownMonthList
	 * @param year
	 */
	
	@SuppressWarnings("rawtypes")
	public static void generateFailedBreakdownByMonthWorkbook(List<BreakdownVO> breakDownMonthList,String year){
		
		try{
			
			String currentWorkBookDate = getDate(7,"ddMMMyyyy");	
			FileInputStream input = new FileInputStream(new File("Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls"));
			HSSFWorkbook workbook = new HSSFWorkbook(input);
			
			HSSFRow row;
			HSSFCell cell;
			BreakdownVO breakdownVO;
			
			
			HSSFSheet monthSheet = workbook.getSheet(year+" - Month");
			
			if(monthSheet == null){
				
				log.info("weeklybreakdown by month sheet unavailable, creating new sheet");
				monthSheet = workbook.createSheet(year+" - Month");
				
				 row  = monthSheet.createRow(0);
				 cell = row.createCell(0);
				 cell.setCellValue("Year");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 cell = row.createCell(1);
				 cell.setCellValue("Month");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 cell = row.createCell(2);
				 cell.setCellValue("Month No");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 cell = row.createCell(3);
				 cell.setCellValue("App Name");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 cell = row.createCell(4);
				 cell.setCellValue("Total Exports");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 cell = row.createCell(5);
				 cell.setCellValue("Failed Exports");
				 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
				 
				workbook.setSheetOrder(year+" - Month", 2);
				
			}
			                 
			Iterator it = breakDownMonthList.iterator();
			
			int column  = 0;
			int CurrentRowNo = 0;
			
			while(it.hasNext()){
				breakdownVO = (BreakdownVO)it.next();
			
			int noOfRows = 	monthSheet.getLastRowNum();
			
		  if(CurrentRowNo == 0){
			  
			for (int rowNum =0; rowNum < noOfRows; rowNum++){
				
				int noOfCells = monthSheet.getRow(rowNum).getLastCellNum();
				
				for(int colNum =0; colNum < noOfCells; colNum ++){
					
					cell = monthSheet.getRow(rowNum).getCell(colNum);
					
					if(cell != null){

						if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
						
						if(cell.getNumericCellValue() == Double.parseDouble(breakdownVO.getMonth_no())){
							
							cell = monthSheet.getRow(rowNum).getCell(colNum+1);
							
							  if(cell != null && cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
								  
							      if(cell.getStringCellValue().equalsIgnoreCase("GRS") || cell.getStringCellValue().equalsIgnoreCase("MIS")){
							    	  
								CurrentRowNo = rowNum;
								break;
							   }
							}
						 }
					  }
				   }
				}
				if(CurrentRowNo != 0){
					break;
				}	
			  }
			
			if(CurrentRowNo == 0){
				CurrentRowNo = monthSheet.getLastRowNum()+ 1;
			}
		  }
		  
		     row  = monthSheet.createRow(CurrentRowNo);
			 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
			 cell.setCellValue(Integer.parseInt(breakdownVO.getYear()));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column++;
			 cell = row.createCell(column);
			 cell.setCellValue(breakdownVO.getMonth());
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column++;
			 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
			 cell.setCellValue(Integer.parseInt(breakdownVO.getMonth_no()));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column++;
			 cell = row.createCell(column);
			 cell.setCellValue(breakdownVO.getApp_name());
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column++;
			 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
			 cell.setCellValue(Integer.parseInt(breakdownVO.getTotal_exports()));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column++;
			 cell = row.createCell(column, HSSFCell.CELL_TYPE_NUMERIC);
			 cell.setCellValue(Integer.parseInt(breakdownVO.getFailed_exports()));
			 cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
			 column = 0;
			 CurrentRowNo++;
				}
			log.info("Iterated through the weeklybreakdown by month values and appended to the sheet");
			for (int j=0; j<7; j++){
				   monthSheet.autoSizeColumn(j);
				}
			input.close();	
			
			FileOutputStream fos = new FileOutputStream("Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls");	
		    workbook.write(fos);
		    fos.close();
		    log.info("weeklybreakdown by month sheet generated successfully and written to workbook");
		    System.out.println("generated successfully");
		    
		     String mailBody = "<FONT style = 'font-family:Calibri'>Hello All,<BR><BR>Please find the Weekly Breakdown of Failed Exports Report for week starting "+currentWorkBookDate+"</FONT>";
		     String subject  = "Weekly Breakdown of Failed Exports Report for week starting "+currentWorkBookDate;
		     String filePath = "Weekly_Failed_breakdown_Week_Starting_"+currentWorkBookDate+".xls";
             String toAddress = "shamsheerc@dnb.com";
		     GenerateMail.sendMail( mailBody, subject, filePath, toAddress);
		     log.info("Failed Breakdown report generated and sent successfully");
		  
	}catch (Exception e) {
		e.printStackTrace();
		log.error("Exception in generating weeklybreakown by month sheet", e);
		System.out.println(e.getMessage());
	}
  }

	public static String convertDate(Date date){
		
		DateFormat dateFormat = new SimpleDateFormat("dd-MMMM-yyyy");
		 String convertedDate = "";
		try {
		Calendar cal = Calendar.getInstance();
                 cal.setTime(date);
         convertedDate =  dateFormat.format(cal.getTime());
		} catch (Exception e) {
			e.printStackTrace();
		}
        return convertedDate;
	}
	
	public static String getDate(int num, String format){ 
	 DateFormat dateFormat = new SimpleDateFormat(format);
     Date myDate = new Date(System.currentTimeMillis());
     Calendar cal = Calendar.getInstance();
     cal.setTime(myDate);
     cal.add(Calendar.DATE, -num);
     String WorkBookDate = dateFormat.format(cal.getTime());
	return WorkBookDate;
	}
}
