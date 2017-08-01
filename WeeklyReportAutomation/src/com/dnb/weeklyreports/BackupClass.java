package com.dnb.weeklyreports;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PrintSetup;

public class BackupClass {
	
	
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void generateSMSXMLWorkBook(Map yearlyUsageMap, String startDate,String endDate) {

		DateFormat dateFormat = new SimpleDateFormat("ddMMMyy");
		Date myDate = new Date(System.currentTimeMillis());
		Calendar cal = Calendar.getInstance();
		final String year = String.valueOf(cal.get(Calendar.YEAR));
		cal.setTime(myDate);
        cal.add(Calendar.DATE, -3);
        
        String lastReportDate = getDate(11,"ddMMMyyyy");
        
        String lastGeneratedreport = getDate(11,"dd-MMM");
        
        String time = dateFormat.format(cal.getTime());
        
        HSSFSheet sheet;
        HSSFCell cell;
        HSSFWorkbook workbook = new HSSFWorkbook();
         
	    sheet = workbook.createSheet(year);
		HSSFRow headerRow = sheet.createRow(0);
		HSSFCell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("Request_Type");
        headerCell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
		//HSSFCell headerCell1 = headerRow.createCell(1);
		//headerCell1.setCellValue(endDate);
		/**
		 * 
		 * Trying to find alternate logic
		 */

		Set set = yearlyUsageMap.entrySet();
		Iterator iterator = set.iterator();

		int nextDate = 1;
		while (iterator.hasNext()) {

			Map.Entry me = (Map.Entry) iterator.next();
			String date = (String) me.getKey();
			
			HSSFCell headerCell1 = headerRow.createCell(nextDate);
			headerCell1.setCellValue(date);
			headerCell1.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			System.out.println("Start date in : " + date);
			TreeMap weeklyMap = (TreeMap) me.getValue();
			Set weeklyset = weeklyMap.entrySet();

			Iterator weeklyIterator = weeklyset.iterator();

			int count = 1;
			int cellcount = 0;

			while (weeklyIterator.hasNext()) {
				Map.Entry weeklyEntry = (Map.Entry) weeklyIterator.next();

				System.out.println("key: " + weeklyEntry.getKey() + " value: "
						+ weeklyEntry.getValue());

				String key = (String) weeklyEntry.getKey();
				int value = Integer.parseInt((String) weeklyEntry.getValue());
				HSSFRow row;
				
				if (nextDate == 1) {
					row = sheet.createRow(count);
				} else
					row = sheet.getRow(count);

				if (nextDate == 1) {
					cell = row.createCell(cellcount);
					cell.setCellValue(key);
					cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				}
				if (key != null) {
					cellcount++;
					if (nextDate != 1) {
						cellcount = nextDate;
					}
					cell = row
							.createCell(cellcount, Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(value);
					cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
					cellcount = 0;
				}
				count++;
			}
			nextDate++;
		}

		for (int i=0; i<10; i++){
			   sheet.autoSizeColumn(i);
			}
		
		String file = "Weekly_GRS_usage_Week_Ending_";
		file = file + time + ".xls";

		try {
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close(); 

			FileOutputStream fos = new FileOutputStream("GRStotalUsage_" + year
					+ ".ser");
			ObjectOutputStream oos = new ObjectOutputStream(fos);
			oos.writeObject(yearlyUsageMap);
			System.out.println("total usage map serialized");
			oos.close();
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	public static void generateSMSXMLGOLFWorkBook(Map totalUsageMap,
			String startDate, String endDate) {
		
		DateFormat dateFormat = new SimpleDateFormat("ddMMMyy");
		Date myDate = new Date(System.currentTimeMillis());
		Calendar cal = Calendar.getInstance();
		final String year = String.valueOf(cal.get(Calendar.YEAR));
		cal.setTime(myDate);
        cal.add(Calendar.DATE, -3);
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet(year);
		sheet.setFitToPage(true);
		sheet.setHorizontallyCenter(true);
		PrintSetup printSetup = sheet.getPrintSetup();
		printSetup.setLandscape(true);

		HSSFRow headerRow = sheet.createRow(0);
		headerRow.setHeightInPoints(12.75f);

		

		HSSFCell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("Week");
		headerCell.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));

		//HSSFCell headerCell1 = headerRow.createCell(1);
		//headerCell1.setCellValue(endDate);
		TreeMap yearlyUsageMap = (TreeMap) totalUsageMap.get(year);
		/**
		 * 
		 * Trying to find alternate logic
		 */

		Set set = yearlyUsageMap.entrySet();
		Iterator iterator = set.iterator();

		int nextDate = 1;
		while (iterator.hasNext()) {

			Map.Entry me = (Map.Entry) iterator.next();
			String date = (String) me.getKey();
			
			HSSFCell headerCell1 = headerRow.createCell(nextDate);
			headerCell1.setCellValue(date);
			headerCell1.setCellStyle(GenerateWorkbook.formatWorkBookCells("HeaderStyle", workbook));
			
			System.out.println("Start date in : " + date);
			TreeMap weeklyMap = (TreeMap) me.getValue();
			Set weeklyset = weeklyMap.entrySet();

			Iterator weeklyIterator = weeklyset.iterator();

			int count = 1;
			int cellcount = 0;

			while (weeklyIterator.hasNext()) {
				Map.Entry weeklyEntry = (Map.Entry) weeklyIterator.next();

				System.out.println("key: " + weeklyEntry.getKey() + " value: "
						+ weeklyEntry.getValue());

				String key = (String) weeklyEntry.getKey();
				int value = Integer.parseInt((String) weeklyEntry.getValue());
				HSSFRow row;
				HSSFCell cell;
				if (nextDate == 1) {
					row = sheet.createRow(count);
				} else
					row = sheet.getRow(count);

				if (nextDate == 1) {
					cell = row.createCell(cellcount);
					cell.setCellValue(key);
					cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
				}
				if (key != null) {
					cellcount++;
					if (nextDate != 1) {
						cellcount = nextDate;
					}
					cell = row
							.createCell(cellcount, Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(value);
					cell.setCellStyle(GenerateWorkbook.formatWorkBookCells("cellStyle", workbook));
					cellcount = 0;
				}
				count++;
			}
			nextDate++;
		}
		
		for (int i=0; i<10; i++){
			   sheet.autoSizeColumn(i);
			}
		String file = "Weekly_SMSXMLGOLF_usage_Week_Ending_";

		String time = dateFormat.format(cal.getTime());
		file = file + time + ".xls";

		try {
			
			
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();

			FileOutputStream fos = new FileOutputStream("GOLFtotalUsage_" + year
					+ ".ser");
			ObjectOutputStream oos = new ObjectOutputStream(fos);
			oos.writeObject(totalUsageMap);
			
			System.out.println("total usage map serialized");
			oos.close();
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	@SuppressWarnings({"finally", "rawtypes", "unchecked"})
	public static Map prepareGOLFTotalUsageMap(Map weeklyMap, String startDate,
			String year) {

		Map totalUsageMap = null;
		Map yearlyUsageMap = null;

		try {
			FileInputStream fis = new FileInputStream("GOLFtotalUsage_" + year
					+ ".ser");
			ObjectInputStream ois = new ObjectInputStream(fis);

			if (ois != null) {
				totalUsageMap = (TreeMap) ois.readObject();
				yearlyUsageMap = (TreeMap) totalUsageMap.get(year);

				if (yearlyUsageMap != null) {
					if (yearlyUsageMap.get(startDate) == null)
						yearlyUsageMap.put(startDate, weeklyMap);
					totalUsageMap.remove(year);
					totalUsageMap.put(year, yearlyUsageMap);
				}
			}
			return totalUsageMap;
		} catch (Exception e) {
			System.out
					.println("Exception occured while deserializing total usage object");
			e.printStackTrace();
		} finally {
			if (yearlyUsageMap == null && totalUsageMap == null) {
				totalUsageMap = new TreeMap();
				yearlyUsageMap = new TreeMap();
				yearlyUsageMap.put(startDate, weeklyMap);
				totalUsageMap.put(year, yearlyUsageMap);
				try {
					FileOutputStream fos = new FileOutputStream("GOLFtotalUsage_"
							+ year + ".ser");
					ObjectOutputStream oos = new ObjectOutputStream(fos);
					oos.writeObject(totalUsageMap);
					oos.close();
					fos.close();

					FileInputStream fis = new FileInputStream("GOLFtotalUsage_"
							+ year + ".ser");
					ObjectInputStream ois = new ObjectInputStream(fis);

					if (ois != null) {
						totalUsageMap = (TreeMap) ois.readObject();
						yearlyUsageMap = (TreeMap) totalUsageMap.get(year);

						if (yearlyUsageMap != null) {
							if (yearlyUsageMap.get(startDate) == null)
								yearlyUsageMap.put(startDate, weeklyMap);
							totalUsageMap.remove(year);
							totalUsageMap.put(year, yearlyUsageMap);
						}
					}
					return totalUsageMap;
				} catch (Exception e) {
					System.out
							.println("Exception in serializing new total usage object");
					e.printStackTrace();
				}
			}
			return totalUsageMap;
		}
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
	
	@SuppressWarnings({"finally", "rawtypes", "unchecked"})
	public static Map prepareGRSTotalUsageMap(Map weeklyMap, String startDate,
			String year) {

		Map yearlyUsageMap = null;

		try {
			FileInputStream fis = new FileInputStream("GRStotalUsage_" + year
					+ ".ser");
			ObjectInputStream ois = new ObjectInputStream(fis);

				yearlyUsageMap = (TreeMap) ois.readObject();
				System.out.println(yearlyUsageMap.size());

				if (yearlyUsageMap != null) {
					if (yearlyUsageMap.get(startDate) == null)
						yearlyUsageMap.put(startDate, weeklyMap);
				}
			return yearlyUsageMap;
			
		} catch (Exception e) {
			System.out
					.println("Exception occured while deserializing total usage object");
			e.printStackTrace();
		} finally {
			if (yearlyUsageMap == null ) {
				yearlyUsageMap = new TreeMap();
				yearlyUsageMap.put(startDate, weeklyMap);
				try {
					FileOutputStream fos = new FileOutputStream("GRStotalUsage_"
							+ year + ".ser");
					ObjectOutputStream oos = new ObjectOutputStream(fos);
					oos.writeObject(yearlyUsageMap);
					oos.close();
					fos.close();

					FileInputStream fis = new FileInputStream("GRStotalUsage_"
							+ year + ".ser");
					ObjectInputStream ois = new ObjectInputStream(fis);
					
						yearlyUsageMap = (TreeMap) ois.readObject();

						if (yearlyUsageMap != null) {
							if (yearlyUsageMap.get(startDate) == null)
								yearlyUsageMap.put(startDate, weeklyMap);
						}
					return yearlyUsageMap;
				} catch (Exception e) {
					System.out
							.println("Exception in serializing new total usage object");
					e.printStackTrace();
				}
			}
			return yearlyUsageMap;
		}
	}


}
