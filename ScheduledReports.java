package org.ce.scheduledtask;

import java.util.*;
import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import org.apache.commons.lang3.StringUtils;

import org.ce.model.Assessment;
import org.ce.model.Carrier;
import org.ce.model.Course;
import org.ce.model.Curriculum;
import org.ce.model.Question;
import org.ce.model.UserCourseResult;
import org.ce.model.Location;
import org.ce.model.SavedReport;
import org.ce.model.ScheduledReport;
import org.ce.model.Subcompany;
import org.ce.model.User;
import org.ce.model.UserCourseStatus;
import org.ce.model.UserCourseStatusRank;
import org.ce.model.UserCurriculum;
import org.ce.model.UserInfo;
import org.ce.model.Role;
import org.ce.util.AppContext;
import org.ce.util.CEProperties;
import org.ce.util.DateUtil;
import org.ce.util.FloatComparatorDesc;
import org.ce.util.ObjectComparatorDesc;
import org.ce.util.StringUtil;
import org.ce.service.CarrierManager;
import org.ce.service.CourseManager;
import org.ce.service.CurriculumManager;
import org.ce.service.LocationManager;
import org.ce.service.QuestionManager;
import org.ce.service.RoleManager;
import org.ce.service.SavedReportManager;
import org.ce.service.ScheduledReportManager;
import org.ce.service.SubcompanyManager;
import org.ce.service.EmailManager;
import org.ce.service.MailEngine;
import org.ce.service.LookupManager;
import org.ce.service.UserCourseResultManager;
import org.ce.service.UserCourseStatusRankManager;
import org.ce.service.UserInfoManager;
import org.ce.service.UserManager;
import org.springframework.context.ApplicationContext;
import org.springframework.mail.SimpleMailMessage;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;

import org.ce.util.SendvelocityMail;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.jsoup.Jsoup;

import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.quartz.Job;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;


public class ScheduledReports implements Job {


    public ScheduledReports() {
    }
   
    ApplicationContext ctx = AppContext.getApplicationContext();
    LookupManager lookup_mgr = (LookupManager) ctx.getBean("lookupManager");       
	UserManager user_mgr = (UserManager) ctx.getBean("userManager");         
	UserCourseResultManager ucr_mgr = (UserCourseResultManager) ctx.getBean("userCourseResultManager");         
    LocationManager location_mgr = (LocationManager) ctx.getBean("locationManager");       
	CurriculumManager curriculum_mgr = (CurriculumManager) ctx.getBean("curriculumManager");         
	ScheduledReportManager scheduledReport_mgr = (ScheduledReportManager) ctx.getBean("scheduledReportManager");  
	SavedReportManager savedReport_mgr = (SavedReportManager) ctx.getBean("savedReportManager");  
	SubcompanyManager subcompany_mgr = (SubcompanyManager) ctx.getBean("subcompanyManager");             
	UserInfoManager userInfo_mgr = (UserInfoManager) ctx.getBean("userInfoManager");  
   	EmailManager email_mgr = (EmailManager) ctx.getBean("emailManager");             
    SimpleMailMessage message = (SimpleMailMessage) ctx.getBean("mailMessage");
	MailEngine mailEngine = (MailEngine) ctx.getBean("mailEngine");
	UserCourseStatusRankManager userCourseStatusRank_mgr = (UserCourseStatusRankManager) ctx.getBean("userCourseStatusRankManager");         
   	QuestionManager question_mgr = (QuestionManager) ctx.getBean("questionManager");             
	CourseManager course_mgr = (CourseManager) ctx.getBean("courseManager");         
	CarrierManager carrier_mgr = (CarrierManager) ctx.getBean("carrierManager");         
	RoleManager role_mgr = (RoleManager) ctx.getBean("roleManager");         
	StringBuffer notStartedCourses = new StringBuffer();
	StringBuffer userAssignedCourses_sb = new StringBuffer();
	StringBuffer msg_expired = new StringBuffer();
	StringBuffer msg_imminent = new StringBuffer();
	StringBuffer msg_expiring = new StringBuffer();
	StringBuffer msg_log = new StringBuffer();
   	Locale locale = new Locale("en");
	SendvelocityMail sendvelocityMail=new SendvelocityMail();

	SXSSFRow eRow = null;
	int cellNum = 0;
	Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
   	String sWidth="3400";
   	String mWidth="5000";
   	String lWidth="10000";
   	String xlWidth="15000";


	private void addCell(String cellValue, String cellStyle){
		SXSSFCell cell = eRow.createCell(cellNum++);
		cell.setCellValue(cellValue);
  		cell.setCellStyle(styles.get(cellStyle));
	}

	private void addCell(String cellValue){
		addCell(cellValue, "defaultCell");
	}
	
	private void addUserInfo(String username, String phone){
		UserInfo userInfo = userInfo_mgr.getUserInfoByUsername(username);
		if (userInfo != null) {
			addCell(userInfo.getLicensenumber() != null ? userInfo.getLicensenumber() : "");
			addCell(userInfo.getLicense_prov() != null ? userInfo.getLicense_prov() : "");
			addCell(userInfo.getJob_title() != null ? userInfo.getJob_title() : "");
			addCell(userInfo.getAddress() != null ? userInfo.getAddress() : "");
			addCell(userInfo.getCity() != null ? userInfo.getCity() : "");
			addCell(userInfo.getProvince() != null ? userInfo.getProvince() : "");
			addCell(userInfo.getPostal() != null ? userInfo.getPostal() : "");
			addCell(!phone.equals("") ? phone : (userInfo.getPhone_number() != null ? userInfo.getPhone_number() : ""));
		}
	}

	private List getDates(JSONObject jObj2){
 		List dates = new ArrayList();
		try {
			if(jObj2.has("dateRange") && jObj2.has("endDateRange")){
				String endDateRange = jObj2.getString("endDateRange");
				Calendar endDate = Calendar.getInstance();
				if (endDateRange!=null && !endDateRange.equals("")) {
					endDate =  DateUtil.getEndDateRange(endDateRange, null);
				}
			  	SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
				String endDateStr = dateFormat.format(endDate.getTime());
				dates.add(dateFormat.format(DateUtil.getEndDateRange(jObj2.getString("dateRange"), endDate).getTime()));
				dates.add(endDateStr);
			}else if(jObj2.has("dateRange")){
				String dateRange = jObj2.getString("dateRange");
				if (dateRange!=null && !dateRange.equals("")) {
					dates.add(DateUtil.getDateRange(dateRange));
				}else{
					dates.add("");
				}
				dates.add("9999-99-99");
			}else{
				dates.add("");
				dates.add("9999-99-99");
			}
		
			if (jObj2.has("timezoneOffset")){
				int t = jObj2.getInt("timezoneOffset");
				if (t==0 || dates.get(0).equals("") && dates.get(1).equals("9999-99-99") ){
					return dates;
				}else{
					return DateUtil.getDatesTimeZone((String)dates.get(0), (String)dates.get(1), t);
				}	
			}else{
				return dates;
			}

		} catch (JSONException e) {
			System.out.println("Scheduled Report Batch - getDates: " + e.getMessage());
			dates = new ArrayList();
			dates.add("");
			dates.add("9999-99-99");
			return dates;
		}
	}
	
	private String getAsOfDate(JSONObject jObj2){
		String date = DateUtil.getCurrentDate().substring(0,10);
		try {
			if(jObj2.has("dateRange")){
				String dateRange = jObj2.getString("dateRange");
				if (dateRange!=null && !dateRange.equals("")) {
					date = DateUtil.getDateRange(dateRange);
				}
			}
		
			if (jObj2.has("timezoneOffset")){
				int t = jObj2.getInt("timezoneOffset");
				if (t!=0 && !date.equals(DateUtil.getCurrentDate().substring(0,10)) ){
					date = DateUtil.getDateTimeZone(date, t);
				}	
			}

		} catch (JSONException e) {
			System.out.println("Scheduled Report Batch - getDates: " + e.getMessage());
		}
		return date;
	}

	private String getDateFilter(List dates, Locale locale, Integer date_format){
		ResourceBundle resource = lookup_mgr.getResourceBundle(locale);
		String dateFilter = resource.getString("label.period")+ " : ";
		if (!((String)dates.get(0)).equals("") || !((String)dates.get(1)).equals("9999-99-99")){
			if (((String)dates.get(0)).equals("") ){
				dateFilter = dateFilter+"~ - ";
			}else{
				dateFilter = dateFilter+DateUtil.getDate(((String)dates.get(0)).substring(0,10), date_format)+" - ";
			}
			if (((String)dates.get(1)).equals("9999-99-99") ){
				dateFilter = dateFilter+DateUtil.getDate(DateUtil.getCurrentDate().substring(0,10), date_format);
			}else{
				dateFilter = dateFilter+DateUtil.getDate(((String)dates.get(1)).substring(0,10), date_format);
			}
		}else{
			dateFilter = dateFilter+resource.getString("label.all");
		}
		return dateFilter;
	}

	private String getAsOfDateFilter(String date, Locale locale, Integer date_format){
		ResourceBundle resource = lookup_mgr.getResourceBundle(locale);
		String dateFilter = "";
		if (!date.substring(0,10).equals(DateUtil.getCurrentDate().substring(0,10)) ){
			dateFilter = resource.getString("report.filter.title.asOf") + " " + DateUtil.getDate(date.substring(0,10), date_format);
		}
		return dateFilter;
	}

	private String getNextReport(String date, Integer frequency){
		if (frequency!=null) {
			int f=frequency.intValue();
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
			Calendar nextReport = DateUtil.convertStringToCalendar(date);
			if (f==1){
				nextReport.add(Calendar.DAY_OF_YEAR, 1);
			}else if (f==2){
				nextReport.add(Calendar.DAY_OF_YEAR, 7);
			}else if (f==3){
				nextReport.add(Calendar.MONTH, 1);
			}else if (f==4){
				nextReport.add(Calendar.MONTH, 3);
			}else if (f==5){
				nextReport.add(Calendar.YEAR, 1);
			}
			return dateFormat.format(nextReport.getTime());
		}else{
			return "";
		}	
	}

	private String[][] getPersonalInfoTitle(Locale locale){
		ResourceBundle resource = lookup_mgr.getResourceBundle(locale);
	   	return new String[][]{
		   		{resource.getString("label.license"), "", mWidth},
		   		{resource.getString("label.issueProv"), "", mWidth},
		   		{resource.getString("userForm.userTitle"), "", sWidth},
		   		{resource.getString("label.address"), "", mWidth},
		   		{resource.getString("label.city"), "", sWidth},
		   		{resource.getString("label.provinceOnly"), "", sWidth},
		   		{resource.getString("label.postalOnly"), "", sWidth},
		   		{resource.getString("label.phone"), "", mWidth}
		   	};
	}
	
   	private String[][] getReportTitle(int reportType, String type, Locale locale){
		ResourceBundle resource = lookup_mgr.getResourceBundle(locale);
		if (reportType==1){	
			return new String[][]{
		   		{resource.getString("userForm.firstName"), "", sWidth},
		   		{resource.getString("userForm.lastName"), "", sWidth},
		   		{resource.getString("userForm.username"), "showUsernames", mWidth},
		   		{resource.getString("label.place"), "", lWidth},
		   		{resource.getString("label.module"), "", lWidth},
		   		{resource.getString("label.payrate"), "showPayRate", sWidth},
		   		{resource.getString("label.startDate"), "", sWidth},
		   		{resource.getString("label.startLoc"), "showDetail", lWidth},
		   		{resource.getString("label.lastVisited"), "", sWidth},
		   		{resource.getString("label.lastLoc"), "showDetail", lWidth},
		   		{resource.getString("label.endDate"), "", sWidth},
		   		{resource.getString("label.endLoc"), "showDetail", lWidth},
		   		{resource.getString("label.score"), "", sWidth},
		   		{resource.getString("label.grade"), "", sWidth},
		   		{resource.getString("label.totalTime"), "showDetail", sWidth},
		   		{resource.getString("label.bookmark"), "showDetail", mWidth},
		   		{resource.getString("label.evalTries"), "showDetail", sWidth},
		   		{resource.getString("label.pass"), "showDetail", sWidth},
		   		{Jsoup.parse(resource.getString("label.numQuestion")).text(), "showDetail", sWidth},
		   		{Jsoup.parse(resource.getString("label.numCorrectAnswer")).text(), "showDetail", sWidth}
		   	};
		}else if (reportType==2){	
		   	return new String[][]{
		   		{resource.getString("userForm.firstName"), "", sWidth},
		   		{resource.getString("userForm.lastName"), "", sWidth},
		   		{resource.getString("userForm.username"), "showUsernames", mWidth},
		   		{resource.getString("label.place"), "", lWidth},
		   		{resource.getString("label.program"), "", lWidth},
		   		{" ", "", lWidth}
		   	};
		}else if (reportType==3){	
			if (type.equals("1")){	
			   	return new String[][]{
			   		{resource.getString("label.manager"), "", mWidth},
			   		{resource.getString("label.status"), "", sWidth},
			   		{resource.getString("label.person"), "", mWidth},
			   		{resource.getString("label.module"), "", lWidth}
			   	};
			}else if (type.equals("2")){	
			   	return new String[][]{
		   			{resource.getString("label.place"), "", lWidth},
		   			{resource.getString("label.status"), "", mWidth},
		   			{resource.getString("label.person"), "", mWidth},
		   			{resource.getString("label.module"), "", lWidth}
			   	};
			}else{
			   	return new String[][]{
			   		{resource.getString("label.person"), "", mWidth},
			   		{resource.getString("label.place"), "", lWidth},
			   		{resource.getString("label.status"), "", mWidth},
			   		{resource.getString("label.module"), "", lWidth}
			   	};
			}
		}else if (reportType==4){	
			if (type.equals("1")){	
				return new String[][]{
			   		{resource.getString("label.user"), "", mWidth},
			   		{resource.getString("label.place"), "", lWidth},
			   		{resource.getString("analyticsReport.qScore"), "", mWidth},
			   		{resource.getString("label.gapRating"), "", mWidth}
			   	};
			}else{
			   	return new String[][]{
			   		{resource.getString("label.module"), "", lWidth},
			   		{resource.getString("label.question"), "", lWidth},
			   		{resource.getString("label.responses"), "", sWidth},
			   		{resource.getString("label.wrong"), "", sWidth},
			   		{resource.getString("label.gapRating"), "", sWidth}
			   	};
			}
		}else if (reportType==5){	
			if (type.equals("1")){	
					return new String[][]{
			   		{resource.getString("label.person"), "", mWidth},
			   		{resource.getString("userForm.username"), "showUsernames", mWidth},
			   		{resource.getString("label.manager"), "", mWidth},
			   		{resource.getString("label.place"), "", lWidth},
			   		{resource.getString("label.program"), "", lWidth},
			   		{resource.getString("label.status"), "", sWidth},
			   		{resource.getString("label.assigned"), "", sWidth},
			   		{resource.getString("label.deadline"), "", sWidth}
			   	};
			}else{
				return new String[][]{
			   		{resource.getString("label.person"), "", mWidth},
			   		{resource.getString("userForm.username"), "showUsernames", sWidth},
			   		{resource.getString("label.manager"), "", mWidth},
			   		{resource.getString("label.place"), "", lWidth},
			   		{resource.getString("label.program"), "", lWidth},
			   		{resource.getString("label.module"), "", lWidth},
			   		{resource.getString("label.status"), "", sWidth},
			   		{resource.getString("label.assigned"), "", sWidth},
			   		{resource.getString("label.date"), "", sWidth},
			   		{resource.getString("label.deadline"), "", sWidth}
			   	};
			}	
		}else{	
			return new String[][]{
				{" ","", "800"},
				{resource.getString("label.question")+" / "+resource.getString("label.answer"),"", "30000"},
				{resource.getString("bf.viewSurvey.numResponses"),"", "3400"},
				{resource.getString("label.percentage"),"", "3000"}
			};
		}
   	}	


	public void execute(JobExecutionContext context) throws JobExecutionException {
		
		JobScheduling jobCE=new JobScheduling("Scheduled Reports");
		if (jobCE.hasDone()){
			return;
		}

    
		System.out.println("---------- Scheduled Reports Start ---------- " + new Date());
		
		
    	SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    	SimpleDateFormat dateFormat1= new SimpleDateFormat("yyyy-MM-dd");

    	Calendar today = Calendar.getInstance();
    	String curr_date = dateFormat.format(today.getTime());

		ScheduledReport scheduledReport=null;
		SavedReport savedReport=null;
	   	Map model = new HashMap();
	   	String filename = "Scheduled Report";
	   	String emailTitle="CarriersEdge Scheduled Report";
	   	String emailTemplate="/org/ce/scheduledtask/ScheduledReport.vm";
	   	int titleHeight = 35;
		NumberFormat df2 = NumberFormat.getNumberInstance();
		df2.setMaximumFractionDigits(2);  


	   	SXSSFWorkbook wb = null;
	   	FileOutputStream os = null;
		int rowNum = 0;
		List recipientList = new ArrayList(); 
		User recipient = null;
		Integer roleLevel = null, roleLevel0 = null;
		StringBuffer msg_log = new StringBuffer();

   		try {
 
   			List reports =lookup_mgr.getResults("from ScheduledReport s where s.status!=0 and s.savedReport.status!=0 and s.recipients is not null and s.next_report is not null and s.next_report!='' and s.next_report<='"+curr_date.substring(0,10)+"'");
   			for (int i=0; i < reports.size(); i++){
				scheduledReport = (ScheduledReport) reports.get(i);
   				savedReport = savedReport_mgr.getSavedReport(scheduledReport.getSavedReportId().toString());
   				Long carrierId = savedReport.getCarrierId();
   				Carrier carrier = carrier_mgr.getCarrier(carrierId.toString());
   				Integer date_format = carrier.getDate_format();
   				System.out.println("--- "+savedReport.getName()+" for carrier "+carrierId+" --- "+new Date());
   				msg_log.append("<br />--- "+savedReport.getName()+" for carrier "+carrierId+" --- "+new Date()+"<br />");

   				recipientList = new ArrayList();
		      	JSONObject recipientjObj = new JSONObject(scheduledReport.getRecipients());
   				if (recipientjObj.has("usernames")){
   					JSONArray recipients = recipientjObj.getJSONArray("usernames");
   					if (recipients != null) { 
   						for (int a=0;a<recipients.length();a++){ 
   							User user = user_mgr.getUser(recipients.getString(a));
   							if (user!=null && user.canNotified()){
   								recipientList.add(new Object[]{user.getRoleLevel(),user});
   							}	
   						} 
   					} 
   				}
   				if (recipientjObj.has("roles")){
   					JSONArray recipients = recipientjObj.getJSONArray("roles");
   					if (recipients != null) { 
   						for (int a=0;a<recipients.length();a++){ 
   							List users=lookup_mgr.getResults("select u from User as u left join u.roles as r where u.status!=0 and r.name='"+recipients.getString(a).replaceAll("-","_")+"' and u.carrierId="+carrierId);
   							if (users.size()>0){	
   								for (int b=0;b<users.size();b++){ 
   									User user = (User)users.get(b);
   		   							if (user!=null && user.canNotified()){
   		   								recipientList.add(new Object[]{user.getRoleLevel(),user});
   		   							}	
   								}	
   							} 
   						} 
   					}
   				}

   				ObjectComparatorDesc comparator = new ObjectComparatorDesc();  
   				Collections.sort(recipientList,comparator);

   				roleLevel = null;
   	   			for (int j=0; j < recipientList.size(); j++){
   	   				Object[] recipientObj=(Object[])recipientList.get(j);
   	   				recipient = (User)recipientObj[1];
   	   				roleLevel0=roleLevel;
   	   				roleLevel=(Integer)recipientObj[0];
   	   				locale = recipient.getLocale();
   	   				ResourceBundle resource = lookup_mgr.getResourceBundle(locale);

   	   				if (!(roleLevel!=null && roleLevel0!=null && roleLevel.intValue()>=40 && roleLevel.equals(roleLevel0))){ 
	   	   			   	String allUserStr = resource.getString("nav.people") + " : "+resource.getString("label.all");
	   	   			   	String allCourseStr = resource.getString("nav.programs") + "/"+ resource.getString("label.modules") + " : "+resource.getString("label.all");

   	   					model = new HashMap();
	   	   				wb = new SXSSFWorkbook(1);
	   	   			
		   	   			XSSFCellStyle style =(XSSFCellStyle) wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			   	   		styles.put("defaultCell", style);
		   	   			
		   	   			XSSFFont font = (XSSFFont)wb.createFont();
		   	   			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		   	   			style.setFont(font);
		   	   			styles.put("headerCell", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(XSSFFont.COLOR_RED);
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			styles.put("redCell", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(new XSSFColor(new Color(0x877836)));
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			styles.put("yellowCell", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(new XSSFColor(new Color(0x13734A)));
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			styles.put("greenCell", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(new XSSFColor(Color.WHITE));
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			style.setFillForegroundColor(new XSSFColor(new Color(0x00005B)));
		   	   			style.setFillPattern(CellStyle.SOLID_FOREGROUND);   	   			
		   	   			styles.put("blueBackground", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(new XSSFColor(Color.BLACK));
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			style.setFillForegroundColor(new XSSFColor(new Color(0xCFD3E0)));
		   	   			style.setFillPattern(CellStyle.SOLID_FOREGROUND);   	   			
		   	   			styles.put("grayBackground", style);
		
		   	   			font = (XSSFFont)wb.createFont();
		   	   			font.setColor(new XSSFColor(Color.BLACK));
		   	   			style = (XSSFCellStyle)wb.createCellStyle();
		   	   			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		   	   			style.setFont(font);
		   	   			style.setFillForegroundColor(new XSSFColor(new Color(0xF7F7FB)));
		   	   			style.setFillPattern(CellStyle.SOLID_FOREGROUND);   	   			
		   	   			styles.put("lightGrayBackground", style);
		
		
		   	   			SXSSFSheet sheet = wb.createSheet();
		   	   			PrintSetup printSetup = sheet.getPrintSetup();
		   	   			printSetup.setLandscape(true);
		   	   			sheet.setFitToPage(true);
		   	   			sheet.setHorizontallyCenter(true);
		   	   			
		   	   			SXSSFCell cell = null;
		   	   			
						rowNum=0; cellNum=0;
		
						JSONObject jObj1 = new JSONObject(savedReport.getQuery_string());
						JSONObject jObj2 = jObj1.getJSONObject("params");

						int reportType = 1;
						if (savedReport.getType()!=null){
							reportType=savedReport.getType().intValue();
						}
						
						List params=new ArrayList();
						List userParams = new ArrayList();
						List courseParams = new ArrayList();
						List otherParams = new ArrayList();
						StringBuffer notes = new StringBuffer();
						boolean allUser=true;
						boolean allCourse=true;
						if (jObj2.names().length()>0){
							String key = "";
							String value = "";
							for(int a = 0; a<jObj2.names().length(); a++){
								key = jObj2.names().getString(a);
								value = jObj2.getString(key);
								if (key.equals("userProgress")) {
									if (!value.equals("null") && !value.equals("None")){
										StringBuffer usernames = new StringBuffer();
										StringBuffer roles = new StringBuffer();
										StringBuffer locations = new StringBuffer();
										StringBuffer subcompanies = new StringBuffer();
										StringBuffer managers = new StringBuffer();
										String[] userListItem = null;
										String[] userList = value.split(",");
										if (userList != null && userList.length > 0) {
											for (int b = 0; b < userList.length; b++){
												userListItem = userList[b].split("-",2);
												if (userListItem[0].equals("user")) {
													User user = user_mgr.getUser(userListItem[1]);
													if (usernames.length()>0) {
														usernames.append(", ");
													}
													usernames.append(user.getFullName());
												} else if (userListItem[0].equals("role")) {
													Role role = role_mgr.getRole(userListItem[1].replaceAll("-","_"));
													if (roles.length()>0) {
														roles.append(", ");
													}
													roles.append(role.getDescription());
												} else if (userListItem[0].equals("place")) {
													Location location = location_mgr.getLocation(userListItem[1]);
													if (locations.length()>0) {
														locations.append(", ");
													}
													locations.append(location.getName());
												} else if (userListItem[0].equals("places")) {
													Subcompany subcompany = subcompany_mgr.getSubcompany(userListItem[1]);
													if (subcompanies.length()>0) {
														subcompanies.append(", ");
													}
													subcompanies.append(subcompany.getName());
												} else if (userListItem[0].equals("manager")) {
													User user = user_mgr.getUser(userListItem[1]);
													if (managers.length()>0) {
														managers.append(", ");
													}
													managers.append(user.getFullName());
												}
											}
										}
										if (usernames.length()>0){
											userParams.add(" " + resource.getString("nav.people") + " : " + usernames);
											allUser=false;
										}
										if (roles.length()>0){
											userParams.add(" " + resource.getString("label.roles") + " : " + roles);
											allUser=false;
										}
										if (locations.length()>0){
											userParams.add(" " + resource.getString("label.locations") + " : " + locations);
											allUser=false;
										}
										if (subcompanies.length()>0){
											userParams.add(" " + resource.getString("label.subcompanies") + " : " + subcompanies);
											allUser=false;
										}
										if (managers.length()>0){
											userParams.add(" " + resource.getString("label.managers") + " : " + managers);
											allUser=false;
										}
									}
								}else if (key.equals("programsProgress")){
									if(!value.equals("null") && !value.equals("None")){
										StringBuffer programs = new StringBuffer();
										StringBuffer modules = new StringBuffer();
										String[] courseListItem = null;
										String[] courseList = value.split(",");
										if (courseList != null && courseList.length > 0) {
											for (int b = 0; b < courseList.length; b++){
												courseListItem = courseList[b].split("-");
												if (courseListItem[0].equals("program")) {
													Curriculum curriculum = curriculum_mgr.getCurriculum(courseListItem[1]);
													if (programs.length()>0) {
														programs.append(", ");
													}
													programs.append(curriculum.getName());
												} else if (courseListItem[0].equals("module")) {
													if (modules.length()>0) {
														modules.append(", ");
													}
													modules.append(lookup_mgr.getCourseName(carrierId, new Long(courseListItem[1]), locale));
												}
											}
										}
										if (programs.length()>0){
											courseParams.add(" " + resource.getString("nav.programs") + " : " + programs);
											allCourse=false;
										}
										if (modules.length()>0){
											courseParams.add(" " + resource.getString("label.modules") + " : " + modules);
											allCourse=false;
										}
									}	
								}else if (key.equals("showDetail") && value!=null && value.equals("true")){
									otherParams.add(resource.getString("report."+key));
								}else if (new String("timezoneOffset,carrierListStr,carrierQueryString,starts,calStart,calEnd,complianceType,analyticsType,carriersAll,statusType,showAll,dateRange,endDateRange,analyticsType,").indexOf(key+",")<0){
									try {
										if (resource.getString("report.filter."+key)!=null && value!=null && value.equals("true") && !(reportType==6 && key.equals("showUsernames"))){
											otherParams.add(resource.getString("report.filter."+key));
										}
									} catch (Exception e) {
										System.out.println("Scheduled Report - "+key+" - "+value+" : "+e.getMessage());
									}
								}	
							}
						}
						
		   				
		   				if (reportType==1){	
		   		   			
			   				Map payrates = new HashMap();
							boolean showPayRate=false;
							List list = lookup_mgr.getResults("select count(*) from CarrierCourseAttribute where status!=0 and pay_rate>0 and carrierId="+carrierId);  
						    if (list!=null && ((Long)list.get(0)).intValue()>0){
		   		   				showPayRate=true;
		   		   			}
							jObj2.put("showPayRate",showPayRate);
		
		
		   		   			boolean showDetail=false;
		   		   			if (jObj2.has("showDetail") && jObj2.getBoolean("showDetail")) {
		   		   				showDetail=true;
		   		   			}
		
	   		   				if (allCourse){
		   		   				params.add(allCourseStr);
	   		   				}else{
	   		   					params.addAll(courseParams);
	   		   				}
	   		   				if (allUser){
		   		   				params.add(allUserStr);
	   		   				}else{
	   		   					params.addAll(userParams);
	   		   				}
			   		 		List dates = getDates(jObj2);
	   		   				params.add(getDateFilter(dates,locale,date_format));
	   		   				params.addAll(otherParams);

	   		   				//title
	   		   				String[][] reportTitleExcel = getReportTitle(reportType,"",locale);
	   		   				String[][] personalInfoTitle = getPersonalInfoTitle(locale);
							eRow = sheet.createRow(rowNum++);
	  	   					cellNum=0;
		  	   				eRow.setHeightInPoints(titleHeight);
				   			for (int a=0; a < reportTitleExcel.length; a++){
								if (reportTitleExcel[a][1].equals("") || !reportTitleExcel[a][1].equals("") && jObj2.has(reportTitleExcel[a][1]) && jObj2.getBoolean(reportTitleExcel[a][1])) {
									addCell(reportTitleExcel[a][0],"headerCell");
			  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
								}
				   			}						
				   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")){
					   			for (int a=0; a < personalInfoTitle.length; a++){
									addCell(personalInfoTitle[a][0],"headerCell");
			  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(personalInfoTitle[a][2]));
					   			}						
				   			}
				   			
				   			boolean showCompleted = false;
		   		   			if (jObj2.has("showCompleted") && jObj2.getBoolean("showCompleted")) {
		   		   				showCompleted = true;
		   		   			}

				   			boolean showTests = false;
		   		   			if (jObj2.has("showTests") && jObj2.getBoolean("showTests")) {
		   		   				showTests = true;
		   		   			}

		   		   			if (jObj1.has("activityQuery")) {
					   			list = lookup_mgr.getResults(jObj1.getString("activityQuery").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%userQueryString%%", recipient.getQueryString("ucr.user.")));
			
					   			if (list != null && list.size() > 0){
					   				for (int a = 0; a < list.size(); a++) {
					   					Object[] obj = (Object[])list.get(a);
					   					
					   					cellNum=0;
					   		   			UserCourseResult ucr = (UserCourseResult) obj[0];
					   		   			
					   		   			if (ucr != null) {
					   		   				if ((!showTests || showTests && (ucr.getPercentage()!=null || ucr.getEval_tries()!=null && ucr.getEval_tries()>0)) && (!showCompleted || showCompleted && ucr.getGrade()!=null && (ucr.getGrade().equals("PASS") || ucr.getGrade().equals("COMPLETE"))) ){
						   		   				eRow = sheet.createRow(rowNum++);
								   		   		
												addCell((String)obj[2]);
												addCell((String)obj[3]);
							   		   			if (jObj2.has("showUsernames") && jObj2.getBoolean("showUsernames")) {
													addCell((String)obj[1]);
							   		   			}
												addCell((String)obj[7]);
						   		   			
						   		   				Course driver_course = ucr.getCourse();
												String course_name = lookup_mgr.getCourseName(carrierId, driver_course.getGroupId(), locale);
												boolean course_inactive = false;
												if (driver_course.getStatus() != null && (driver_course.getStatus().intValue() < 21 || driver_course.getStatus().intValue() > 50)){
					  									course_inactive = true;
												}
												if (course_inactive || jObj2.has("showVariations") && jObj2.getBoolean("showVariations")){
													course_name = lookup_mgr.getCourseName(carrierId, driver_course)+" ("+resource.getString("label.inactiveModule")+")";
												}
												addCell(course_name);
							   		   			if (showPayRate) {
													Float payrate = (Float)payrates.get(carrierId+'-'+ucr.getCourseId());
													if (payrate == null) {
														List rates = lookup_mgr.getResults("select pay_rate from CarrierCourseAttribute where status!=0 and courseId="+ucr.getCourseId()+" and carrierId="+carrierId);
														if (rates != null && rates.size() > 0){
															payrate = (Float)rates.get(0);
															payrates.put(carrierId+'-'+ucr.getCourseId(),payrate);
														}
													}
													addCell(payrate != null && payrate > 0 ? "$"+String.format("%.2f", payrate) : "");
							   		   			}
												addCell(ucr.getStart_date() != null && !ucr.getStart_date().equals("")? DateUtil.getDate(ucr.getStart_date(), date_format) : "");
							   		   			if (showDetail) {
													addCell(ucr.getStart_loc() != null ? ucr.getStart_loc() : "");
							   		   			}
				
												addCell(ucr.getLast_access() != null && !ucr.getLast_access().equals("")? DateUtil.getDate(ucr.getLast_access(), date_format) : "");
							   		   			if (showDetail) {
													addCell(ucr.getLast_loc() != null ? ucr.getLast_loc() : "");
							   		   			}
				
												addCell(ucr.getEnd_date() != null && !ucr.getEnd_date().equals("")? DateUtil.getDate(ucr.getEnd_date(), date_format) : "");
							   		   			if (showDetail) {
													addCell(ucr.getEnd_loc() != null ? ucr.getEnd_loc() : "");
							   		   			}
							   		   			
					   							String grade = ucr.getGrade() != null ? ucr.getGrade() : "";
							   		   			if (grade.equals("PASS") || grade.equals("COMPLETE")){
													addCell((ucr.getPercentage() != null ? df2.format(ucr.getPercentage())+"%" : ""),"greenCell");
							   		   			}else if (grade.equals("FAIL")){
													addCell((ucr.getPercentage() != null ? df2.format(ucr.getPercentage())+"%" : ""),"redCell");
							   		   			}else{
													addCell((ucr.getPercentage() != null ? df2.format(ucr.getPercentage())+"%" : ""));
							   		   			}
				
							   		   			if (grade.equals("PASS") || grade.equals("COMPLETE")){
													addCell(grade,"greenCell");
							   		   			}else if (grade.equals("FAIL")){
													addCell(grade,"redCell");
							   		   			}else{
													addCell(grade);
							   		   			}
				
							   		   			if (showDetail) {
													String totalTime = "";
													if (ucr.getTotal_time() != null && ucr.getTotal_time().intValue() > 0){
														int t = (ucr.getTotal_time()).intValue();
														totalTime = t/3600000+":";
														totalTime = totalTime+((t%3600000)/60000<10?"0"+(t%3600000)/60000:""+(t%3600000)/60000)+":";
														totalTime = totalTime+((t%60000)/1000<10?"0"+(t%60000)/1000:""+(t%60000)/1000);
													}
													addCell(totalTime);
													addCell(ucr.getBookmark() != null ? ucr.getBookmark() : "");
													addCell(ucr.getEval_tries() != null ? ucr.getEval_tries().toString() : "0");
													addCell(ucr.getPass() != null ? ucr.getPass().toString() : "");
													addCell(ucr.getNum_questions() != null ? ucr.getNum_questions().toString() : "");
													addCell(ucr.getNum_correct_answers() != null ? ucr.getNum_correct_answers().toString() : "");
							   		   			}
				
							   		   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")) {
							   		   				addUserInfo(ucr.getUsername(),obj[8].toString());
					   							}
						   		   			}
						   				}	
						   			}
				  		   		}
		   		   			}
		  		   			if (jObj2.has("showNoData") && jObj2.getBoolean("showNoData") && jObj1.has("activityQueryNoData") && !showTests && !showCompleted) {
					   			list = lookup_mgr.getResults(jObj1.getString("activityQueryNoData").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%userQueryString%%", recipient.getQueryString()));
					   			if (list != null && list.size() > 0) {	
					   				for (int a = 0; a < list.size(); a++) {
					   					Object[] obj = (Object[])list.get(a);
					   					cellNum=0;
					   					eRow = sheet.createRow(rowNum++);
										addCell((String)obj[1]);
										addCell((String)obj[2]);
					   		   			if (jObj2.has("showUsernames") && ((Boolean)jObj2.get("showUsernames")).booleanValue()) {
											addCell((String)obj[0]);
					   		   			}
										addCell((String)obj[6]);
					   		   			if (jObj2.has("showPersonalInfo") && ((Boolean)jObj2.get("showPersonalInfo")).booleanValue()) {
					   		   				int num= 0;
					   		   				if (showDetail){
					   		   					num=15;
					   		   				}else{
					   		   					num=6;
					   		   				}
					   		   				if (showPayRate){
					   		   					num++;
					   		   				}	
				   		   					cellNum=cellNum+num;
					   		   				addUserInfo((String)obj[0],"");
			   							}
					   				}
					   			}
		  		   			}	
		
		  		   			
		   				}else if (reportType==2){	
		  	   		   		
	   		   				if (allCourse){
		   		   				params.add(allCourseStr);
	   		   				}else{
	   		   					params.addAll(courseParams);
	   		   				}
	   		   				if (allUser){
		   		   				params.add(allUserStr);
	   		   				}else{
	   		   					params.addAll(userParams);
	   		   				}

		   					String asOfDate=getAsOfDate(jObj2);
			   		   		if (!getAsOfDateFilter(asOfDate,locale,date_format).equals("")){
			   		   			params.add(getAsOfDateFilter(asOfDate,locale,date_format));
			   		   		}
	   		   				params.addAll(otherParams);

	   		   				String[][] reportTitleExcel = getReportTitle(reportType,"",locale);
							eRow = sheet.createRow(rowNum++);
							cellNum = 0;
		  	   				eRow.setHeightInPoints(titleHeight);
							for (int a=0; a < reportTitleExcel.length; a++){
								if (reportTitleExcel[a][1].equals("") || !reportTitleExcel[a][1].equals("") && jObj2.has(reportTitleExcel[a][1]) && jObj2.getBoolean(reportTitleExcel[a][1])) {
									addCell(reportTitleExcel[a][0],"headerCell");
			  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
								}
				   			}						
		
		  		   			if (jObj1.has("ProgressQueryUCs")) {
					   			List list = lookup_mgr.getResults(jObj1.getString("ProgressQueryUCs").replaceAll("%%asOfDate%%", asOfDate).replaceAll("%%userQueryString%%", recipient.getQueryString("uc.user.")));
					   			if (list != null && list.size() > 0){
					   				int status = 0,	status1 = 0, status2 = 0, status3 = 0, maxCourseNum = 0, complete = 0, inprogress = 0, notstarted = 0;
					   				String cellStyle="";
					   				for (int a = 0; a < list.size(); a++) {
		  				   					
						   				UserCurriculum uc = (UserCurriculum)list.get(a);
			  				   			User user = user_mgr.getUser(uc.getUsername());
			  				   			Location location = location_mgr.getLocation(user.getLocationId().toString());
			  				   			Curriculum curriculum = curriculum_mgr.getCurriculum(uc.getCurriculumId().toString());
						   				String locName = location.getName();
						   				if (location.getSubcompanyId() != null ) {
							   				Subcompany subcompany = subcompany_mgr.getSubcompany(location.getSubcompanyId().toString());
						   					locName = subcompany.getName()+" - "+locName;
						   				}
					   					cellNum=0;
			
					   					eRow = sheet.createRow(rowNum++);
							   		   		
										addCell(user.getFirstName());
										addCell(user.getLastName());
					   		   			if (jObj2.has("showUsernames") && jObj2.getBoolean("showUsernames")) {
											addCell(user.getUsername());
					   		   			}
										addCell(locName);
										addCell(curriculum.getName(locale));
											
			  				   			status = 0;
			  				   			status1 = 0;
			  				   			status2 = 0;
			  				   			status3 = 0;
			
			  				   			for (int b = 0; b < 3; b++){
			  				   				List courses = null;
			  				   				if (b == 0){
			  				   					courses = lookup_mgr.getResults(jObj1.getString("ProgressQueryStatus1").replaceAll("%%username%%", uc.getUsername()).replaceAll("%%curriculumId%%", uc.getCurriculumId().toString()).replaceAll("%%asOfDate%%", asOfDate));
		  				   						cellStyle="greenCell";
			  				   					if (courses != null && courses.size() > 0 && status < 1){
			  				   						status = 1;
			  				   					}
			  				   				} else if (b == 1) {
			  				   					courses = lookup_mgr.getResults(jObj1.getString("ProgressQueryStatus2").replaceAll("%%username%%", uc.getUsername()).replaceAll("%%curriculumId%%", uc.getCurriculumId().toString()).replaceAll("%%asOfDate%%", asOfDate));
		  				   						cellStyle="yellowCell";
			  				   					if (courses != null && courses.size() > 0 && status < 2){
			  				   						status = 2;
			  				   					}
			  				   				} else {
			  				   					courses = lookup_mgr.getResults(jObj1.getString("ProgressQueryStatus3").replaceAll("%%username%%", uc.getUsername()).replaceAll("%%curriculumId%%", uc.getCurriculumId().toString()).replaceAll("%%asOfDate%%", asOfDate));
		  				   						cellStyle="redCell";
			  				   					if (courses != null && courses.size() > 0 && status < 3){
			  				   						status = 3;
			  				   					}
			  				   				}
			  				   				if (courses != null){
			  				   					if (maxCourseNum < courses.size()) {
			  				   						maxCourseNum = courses.size();
			  				   					}
			  				   					for (int c = 0; c < courses.size(); c++){
			  				   						Long groupId = (Long)courses.get(c);
			  										addCell(lookup_mgr.getCourseName1(user, groupId, locale), cellStyle);
			  				  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(lWidth));
			  				   						if (status == 1) {
			  				   							status1++;
			  				   						} else if (status == 2) {
			  				   							status2++;
			  				   						} else if (status == 3) {
			  				   							status3++;
			  				   						}
			  				   					}
			  				   				}
			  				   			}
			  				   			if (status == 1) {
			  				   				complete += 1;
			  				   			} else if (status == 2) {
			  				   				inprogress += 1;
			  				   			} else if (status == 3) {
			  				   				notstarted += 1;
			  				   			}
			  		  		   		}
			   		   				notes.append(resource.getString("progressReport.summary") + " : " + 
			   		   								"<font color='13734A'>"+resource.getString("label.current")+": "+df2.format(((float)complete/(complete+inprogress+notstarted)*10000)/100f)+"%; </font>"+
			   		   								"<font color='877836'>"+resource.getString("ucsd.action.continue")+": "+df2.format(((float)inprogress/(complete+inprogress+notstarted)*10000)/100f)+"%; </font>"+
			   		   								"<font color='red'>"+resource.getString("ucsd.status.notstarted")+", "+resource.getString("ucsd.status.incomplete")+", "+resource.getString("ucsd.status.failed")+": "+df2.format(((float)notstarted/(complete+inprogress+notstarted)*10000)/100f)+"%"+"</font>");
		
						   		}
			
			   				}
		   				
		   				
				   			
		   				}else if (reportType==3){	
		  	   		   			
		   					String cellStyle = "";
		   					
		   					String complianceType="1";
		   		   			if (jObj2.has("complianceType") ) {
		   		   				complianceType=jObj2.getString("complianceType");
		   		   				params.add(resource.getString("header.report.compliance")+" "+resource.getString("label.type")+" : "+resource.getString("complianceReport.type"+complianceType));
		   		   			}
		
	   		   				String[][] reportTitleExcel = getReportTitle(reportType,complianceType,locale);
	   		   				String[][] personalInfoTitle = getPersonalInfoTitle(locale);
							eRow = sheet.createRow(rowNum++);
		   					cellNum=0;
		  	   				eRow.setHeightInPoints(titleHeight);
				   			for (int a=0; a < reportTitleExcel.length; a++){
								addCell(reportTitleExcel[a][0],"headerCell");
		  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
				   			}						
							if (!complianceType.equals("1") && !complianceType.equals("2")){
		   		   				if (allUser){
			   		   				params.add(allUserStr);
		   		   				}else{
		   		   					params.addAll(userParams);
		   		   				}
							}
				   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")){
					   			for (int a=0; a < personalInfoTitle.length; a++){
									addCell(personalInfoTitle[a][0],"headerCell");
			  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(personalInfoTitle[a][2]));
					   			}						
				   			}
		
		   		   			String asOfDate=getAsOfDate(jObj2);
			   		   		if (!getAsOfDateFilter(asOfDate,locale,date_format).equals("")){
			   		   			params.add(getAsOfDateFilter(asOfDate,locale,date_format));
			   		   		}
	   		   				params.addAll(otherParams);
			   			
		  		   			if (jObj1.has("complianceQuery")) {
					   			List list = lookup_mgr.getResults(jObj1.getString("complianceQuery").replaceAll("%%asOfDate%%", asOfDate).replaceAll("%%userQueryString%%", recipient.getQueryString()));
		  		
					   			if (list != null && list.size() > 0){
					   				User user = null, ucsUser = null;
					   				Location location = null;
					   				Subcompany subcompany = null;
					   				String queryString = "", course_name = "";
					   				int status = 0, prevStatus = 0, compliant = 0, workingcompliant = 0, notcompliant = 0, compliantCount = 0;
		
				   					for (int a = 0; a < list.size(); a++){
				   						if (complianceType.equals("1")) {
				   							user = (User) list.get(a);
				   							queryString = jObj1.getString("complianceQueryUCS").replaceAll("%%asOfDate%%", asOfDate).replaceAll("%%userString%%", user.getQueryString("ucs.user.")).replaceAll("%%userQueryString%%", recipient.getQueryString("ucs.user."));
				   						} else if (complianceType.equals("2")){
				   							location = (Location) list.get(a);
				   							queryString = jObj1.getString("complianceQueryUCS").replaceAll("%%asOfDate%%", asOfDate).replaceAll("%%locationId%%", location.getId().toString()).replaceAll("%%userQueryString%%", recipient.getQueryString("ucs.user."));
				   						}else{
				   							user = (User) list.get(a);
				   							queryString = jObj1.getString("complianceQueryUCS").replaceAll("%%asOfDate%%", asOfDate).replaceAll("%%username%%", user.getUsername()).replaceAll("%%userQueryString%%", recipient.getQueryString("ucs.user."));
				   						}
		
			   							List ucss = lookup_mgr.getResults(queryString);
			   							if (ucss != null && ucss.size() > 0){
			   								for (int b = 0; b < ucss.size(); b++){
				   									UserCourseStatus ucs = (UserCourseStatus)ucss.get(b);
				   									if (ucs != null) {
				   										String locName = "";
				   										course_name = "";
				   										UserCourseStatusRank ucsr=userCourseStatusRank_mgr.getUserCourseStatusRank(ucs.getStatusId().toString());
				   										status = ucsr.getStatus_weight().intValue();
				   										ucsUser = user_mgr.getUser(ucs.getUsername());
				   										if (b > 0 && (prevStatus != status || status != 4)) {
				   											break;
				   										} else {
				   											prevStatus = status;
				   										}
				   										if (status != 1) {
				   											course_name = lookup_mgr.getCourseName1(ucsUser, ucs.getCourseId(), locale);
				   										}
				   										if (!complianceType.equals("2")) {
				   											location = location_mgr.getLocation(user.getLocationId().toString());
				   										}
				   										locName = location.getName();
				   										if (location.getSubcompanyId()!=null){
					   										subcompany = subcompany_mgr.getSubcompany(location.getSubcompanyId().toString());
				   											if (subcompany != null && subcompany.getName() != null && !subcompany.getName().equals("")) {
					   											locName = subcompany.getName()+" - "+locName;
					   										}
				   										}
				   										if (b == 0) {
				   											compliantCount++;
				   											if (status == 1) {
				   												compliant++;
						  				   						cellStyle="greenCell";
				   											} else if (status == 5) {
				   												notcompliant++;
						  				   						cellStyle="redCell";
				   											} else if (status > 1 && status < 5) {
				   												workingcompliant++;
						  				   						cellStyle="yellowCell";
				   											}
				   										}	
				   										eRow = sheet.createRow(rowNum++);
				   					   					cellNum=0;
			   											if (complianceType.equals("1")) {
					   										if (b == 0) {
				   												addCell(user.getFullName());
				   												addCell(ucsr.getStatus_name(), cellStyle);
					   										}else{
					   											cellNum = cellNum+2;
					   										}
				   											addCell(ucsUser.getFullName());
					   										addCell(course_name);
					   					   		   			if (b==0 && jObj2.has("showPersonalInfo") && ((Boolean)jObj2.get("showPersonalInfo")).booleanValue()) {
					   					   		   				addUserInfo(user.getUsername(),"");
					   			   							}
			   											}else if (complianceType.equals("2")) {
					   										if (b == 0) {
				   												addCell(locName);
				   												addCell(ucsr.getStatus_name(), cellStyle);
					   										}else{
					   											cellNum = cellNum+2;
					   										}
					   										addCell(ucsUser.getFullName());
					   										addCell(course_name);
			   											} else {
					   										if (b == 0) {
				   												addCell(user.getFullName());
						   										addCell(locName);
				   												addCell(ucsr.getStatus_name(), cellStyle);
					   										}else{
					   											cellNum = cellNum+3;
					   										}
					   										addCell(course_name);
					   					   		   			if (b==0 && jObj2.has("showPersonalInfo") && ((Boolean)jObj2.get("showPersonalInfo")).booleanValue()) {
					   					   		   				addUserInfo(user.getUsername(),"");
					   			   							}
			   											}
				   									}
				   								}
				   							}
				   						}
			   							notes.append(resource.getString("complianceReport.summary") + " : " + 
			   								"<font color='red'>"+resource.getString("snapshot.usercompliance.noncompliant")+": "+(compliant+notcompliant+workingcompliant>0?df2.format(((float)notcompliant/(compliant+notcompliant+workingcompliant)*10000)/100f)+"%;    ":"    ")+"</font>"+
			   								"<font color='877836'>"+resource.getString("snapshot.usercompliance.towardscompliant")+": "+(compliant+notcompliant+workingcompliant>0?df2.format(((float)workingcompliant/(compliant+notcompliant+workingcompliant)*10000)/100f)+"%;    ":"    ")+"</font>"+
			   								"<font color='13734A'>"+resource.getString("snapshot.usercompliance.compliant")+": "+(compliant+notcompliant+workingcompliant>0?df2.format(((float)compliant/(compliant+notcompliant+workingcompliant)*10000)/100f)+"%":"")+"</font>");
				   					}
		  		   				}
					   			
		  	   				}else if (reportType==4){	

			  	   				NumberFormat df = NumberFormat.getNumberInstance();
				  	   			df.setMaximumFractionDigits(2);
				  	   			df.setMinimumFractionDigits(2);
				  	   			df.setGroupingUsed(false);
		  	  	   		   			
		  	   		   			String analyticsType="2";
		  	   		   			if (jObj2.has("analyticsType") ) {
		  	   		   				analyticsType=jObj2.getString("analyticsType");
		  	   		   				params.add(resource.getString("analyticsReport.title")+" "+resource.getString("label.type")+" : "+resource.getString("analyticsReport.type"+analyticsType));
		  	   		   			}
		
							    String userPrefix = "";
							    if (analyticsType.equals("1")) {
							    	userPrefix = "u.";
							    } else {
							    	userPrefix = "aua.user.";
							    }

		  						boolean showDetail=false;
		  	   		   			if (jObj2.has("showDetail") && jObj2.getBoolean("showDetail")) {
		  	   		   				showDetail=true;
		  	   		   			}
		
		   		   				if (allCourse){
			   		   				params.add(allCourseStr);
		   		   				}else{
		   		   					params.addAll(courseParams);
		   		   				}
		   		   				if (allUser){
			   		   				params.add(allUserStr);
		   		   				}else{
		   		   					params.addAll(userParams);
		   		   				}
	  	   		   				
				   		 		List dates = getDates(jObj2);
	  	   		   				params.add(getDateFilter(dates,locale,date_format));
		   		   				params.addAll(otherParams);

	  	   		   				List listSort=new ArrayList();
		  						if (analyticsType.equals("1")) {
		  							notes.append("<span style='color:red'>"+resource.getString("analyticsReport.userGaps.critical")+"</span> - "+resource.getString("analyticsReport.userGaps.critical.text")+"<br />"+
		  									resource.getString("analyticsReport.userGaps.low")+" - "+resource.getString("analyticsReport.userGaps.low.text")+"<br />"+
		  									resource.getString("analyticsReport.userGaps.high")+" - "+resource.getString("analyticsReport.userGaps.high.text")+"<br />"+
		  									"<span style='color:#13734A'>"+resource.getString("analyticsReport.userGaps.star")+"</span> - "+resource.getString("analyticsReport.userGaps.star.text"));
		  				   			List carrier_drivers = lookup_mgr.getResults(jObj1.getString("getAnalyticsDrivers").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%userQueryString%%", recipient.getQueryString(userPrefix)));
		  							if (carrier_drivers != null){
		  								for (int a = 0; a < carrier_drivers.size(); a++){
		  									Object[] userObj = (Object[])carrier_drivers.get(a);
		  									if (userObj != null){
		  										String username = userObj[0].toString();
		  										List userCourseResults = lookup_mgr.getResults(jObj1.getString("getAnalyticsUCRs").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%username%%", username));
		  										Object[] ucrObj = (Object[])userCourseResults.get(0);
		  										if (ucrObj != null) {
		  											float Qscore = (float)0;
		  											float aggregate_score = ((Double)ucrObj[1]).floatValue();
		  											float gapRating = (float)1;
		  											int numCourses = ((Long)ucrObj[0]).intValue();
		  											if (numCourses > 0){
		  												Qscore = (float)((aggregate_score/(float)numCourses)*10);
		  												gapRating = 1000 - Qscore;
		  												listSort.add(new Object[]{df.format(gapRating),userObj[1].toString()+" "+userObj[2].toString(),userObj[6].toString(),df.format(Qscore)});
		  											}
		  										}
		  									}
		  								}
		  							}
		  						} else {
		  							notes.append(" <span style='color:red'>"+resource.getString("analyticsReport.userGaps.critical")+"</span> - "+resource.getString("analyticsReport.fleetGaps.critical.text")+"<br />"+
		  									resource.getString("analyticsReport.userGaps.low")+" - "+resource.getString("analyticsReport.fleetGaps.low.text")+"<br />"+
		  									resource.getString("analyticsReport.userGaps.high")+" - "+resource.getString("analyticsReport.fleetGaps.high.text")+"<br />"+
											"<span style='color:#13734A'>"+resource.getString("analyticsReport.userGaps.star")+"</span> - "+resource.getString("analyticsReport.fleetGaps.star.text"));
		  							String allCarrierQs = jObj1.getString("getAnalyticsQuestions").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%userQueryString%%", recipient.getQueryString(userPrefix));
		  							if (allCarrierQs != null ) {
		  								List carrier_questions = lookup_mgr.getResults(allCarrierQs);
		  								if (carrier_questions != null) {
		  									for (int a = 0; a < carrier_questions.size(); a++) {
		  										Object[] qObj = (Object[])carrier_questions.get(a);
		  										if (qObj != null) {
		  											Long qId = (Long)qObj[0];
		  											Long aId = (Long)qObj[1];
		  											List list = lookup_mgr.getResults("select a.id from Answer a where a.questionId="+qId+" and a.correct=1");
		  											Long corr = null;
		  											if (list != null && list.size() > 0) {
		  												corr = (Long)list.get(0);
		  												if (corr != null){
		  													String courseName = "";
		  													Long gId = null;
		  													list = lookup_mgr.getResults("from Course c where c.id in (select a.courseId from Assessment a where a.id="+aId+")");
		  													if (list != null && list.size() > 0) {
		  														Course course = (Course)list.get(0);
		  														if (course != null) {
		  															gId = course.getGroupId();
		  															if (course.getStatus() != null && (course.getStatus().intValue() < 21 || course.getStatus().intValue() > 50)){
		  																courseName = lookup_mgr.getCourseName(carrierId, course)+" ("+resource.getString("label.inactiveModule")+")";
		  															} else {
		  																courseName = lookup_mgr.getCourseName(carrierId, gId, locale);
		  															}
		  														}
		  													}
		
		  													List allAUAs = lookup_mgr.getResults(jObj1.getString("getAnalyticsAUAs").replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1)).replaceAll("%%corr%%", corr.toString()).replaceAll("%%qId%%", qId.toString()).replaceAll("%%userQueryString%%", recipient.getQueryString(userPrefix)));
		  													Object[] allAUAObj = (Object[])allAUAs.get(0);
		  													if (allAUAObj != null) {
		  														int num_responses = ((Long)allAUAObj[0]).intValue();
		  														int num_wrong = ((Long)allAUAObj[1]).intValue();
		  														if (num_responses > 0 && num_wrong > 0) {
		  															float gap = (float)((float)num_wrong/(float)num_responses)*1000;
		  			  												listSort.add(new Object[]{df.format(gap),courseName.toString(),qObj[2].toString(),String.valueOf(num_responses),String.valueOf(num_wrong)});
		  														}
		  													}
		  												}
		  											}
		  										}
		  									}
		  								}
		  							}
		  						}
		
		   		   				String[][] reportTitleExcel = getReportTitle(reportType,analyticsType,locale);
		  						eRow = sheet.createRow(rowNum++);
		  	   					cellNum=0;
			  	   				eRow.setHeightInPoints(titleHeight);
	  				   			for (int a=0; a < reportTitleExcel.length; a++){
	  								addCell(reportTitleExcel[a][0],"headerCell");
			  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
	  				   			}						
		
								if (listSort!=null && listSort.size()>0){
									FloatComparatorDesc fcomparator = new FloatComparatorDesc();  
						    		Collections.sort(listSort,fcomparator);
						    		
						    		String cellStyle0 = "redCell";
						    		String cellStyle = "";
						    		float gap0 = -1;
						    		float gap = 0;
						    		int totalRow=listSort.size();
						    		int c1 = (int)(totalRow*0.2);
						    		int c2 = (int)(totalRow*0.6) + c1;
		
									for (int a = 0; a < listSort.size(); a++){
										Object[] obj = (Object[])listSort.get(a);
										eRow = sheet.createRow(rowNum++);
					  	   				cellNum=0;
					  	   				addCell(obj[1].toString());
					  	   				addCell(Jsoup.parse(StringEscapeUtils.unescapeHtml3(obj[2].toString())).text());
					  	   				addCell(obj[3].toString());
				  						if (!analyticsType.equals("1")) {
						  	   				addCell(obj[4].toString());
				  						}
		
		    							cellStyle0 = cellStyle;
		    							gap0 = gap;
				  						gap = Float.valueOf((String)obj[0]).floatValue();
				    					if (a >= c2){
				    						if (gap == gap0){
				    							cellStyle = cellStyle0;
				    						}else{
				    							cellStyle = "greenCell";
				    						}
				    					} else{
				    						if (a >= c1){	
					    						if (gap == gap0){
					    							cellStyle = cellStyle0;
					    						}else{
					    							cellStyle = "";
					    						}
				    						} else {
				    							cellStyle = "redCell";
				    						}
				    					} 
				    					if (gap == 0){
			    							cellStyle = "greenCell";
				    					}
				  						
			  							addCell(obj[0].toString(),cellStyle);
									}
								}	
						 		
						   			
		  	   				}else if (reportType==5){	
		  	   					
		  	   					int rptTotalRecords = 0, totalPrograms = 0, timezoneOffset = 0, totalComplete = 0, inProgress = 0, notStarted = 0, pTotalComplete = 0, pInProgress = 0, pNotStarted = 0;
		  	   					String courseStatus = "";
		  	   				
		  	   		   			String statusType="1";
		  	   		   			if (jObj2.has("statusType") ) {
		  	   		   				statusType=jObj2.getString("statusType");
		  	   		   			}
	  	   		   				params.add(resource.getString("statusReport.title")+" "+resource.getString("label.type")+" : "+resource.getString("statusReport.type"+statusType));
		
		  						boolean showIncompleted=false;
		  	   		   			if (jObj2.has("showIncompleted") && jObj2.getBoolean("showIncompleted")) {
		  	   		   				showIncompleted=true;
		  	   		   			}
		
		  	   		   			boolean includeIndividualModules=false;
		  	   		   			if (statusType.equals("2") && jObj2.has("includeIndividualModules") && jObj2.getBoolean("includeIndividualModules")) {
		  	   		   				includeIndividualModules=true;
		  	   		   			}
		  	   		   			
		   		   				if (allCourse){
			   		   				params.add(allCourseStr);
		   		   				}else{
		   		   					params.addAll(courseParams);
		   		   				}
		   		   				if (allUser){
			   		   				params.add(allUserStr);
		   		   				}else{
		   		   					params.addAll(userParams);
		   		   				}

		   		   				String asOfDate=getAsOfDate(jObj2);
				   		   		if (!getAsOfDateFilter(asOfDate,locale,date_format).equals("")){
				   		   			params.add(getAsOfDateFilter(asOfDate,locale,date_format));
				   		   		}
		   		   				params.addAll(otherParams);

		   		   				Calendar startDateCal=DateUtil.convertStringToCalendar(asOfDate);
		
		   		   				String[][] reportTitleExcel = getReportTitle(reportType,statusType,locale);
		   		   				String[][] personalInfoTitle = getPersonalInfoTitle(locale);
		  						eRow = sheet.createRow(rowNum++);
		  	   					cellNum=0;
			  	   				eRow.setHeightInPoints(titleHeight);
	  				   			for (int a=0; a < reportTitleExcel.length; a++){
	  								if (reportTitleExcel[a][1].equals("") || !reportTitleExcel[a][1].equals("") && jObj2.has(reportTitleExcel[a][1]) && jObj2.getBoolean(reportTitleExcel[a][1])) {
	  									addCell(reportTitleExcel[a][0],"headerCell");
				  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
	  								}
	  				   			}						
					   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")){
						   			for (int a=0; a < personalInfoTitle.length; a++){
										addCell(personalInfoTitle[a][0],"headerCell");
				  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(personalInfoTitle[a][2]));
						   			}						
					   			}
					   			
					   			boolean summarizeByPlace = false;
					   			String lastRegionProgram = "";
					   			JSONObject jObj3 = new JSONObject();
					 
					   			summarizeByPlace = jObj2.getBoolean("summarizeByPlace");
					   			if(summarizeByPlace){
					   				List listByPlace = lookup_mgr.getResults(jObj1.getString("getStatuQueryPlace").replaceAll("%%userQueryString%%", recipient.getQueryString("uc.user.")));
					   				if (listByPlace != null && listByPlace.size() > 0){
					   					for (int a = 0; a < listByPlace.size(); a++) {
					   						Object[] ucObj = (Object[])listByPlace.get(a);
					   						if (ucObj != null) {
					   							String programName = ucObj[1].toString();
					   							String regionName = ucObj[6].toString();
					   							String programDeadline = ucObj[8].toString();
					   							String programAssigned = ucObj[9].toString();
					   							String lastActivity = ucObj[11] != null ? ucObj[11].toString() : "";
					   							Long programId = (Long)ucObj[0];
					   							Long locationId = (Long)ucObj[3];
					   							Long regionId = (Long)ucObj[5];
					   							Long numUsers = (Long)ucObj[7];
					   							Long numComplete = (Long)ucObj[10];
					   							double percComplete = Math.round((numComplete.floatValue() / numUsers.floatValue() * 100) * 10) / 10.0;
					   							DecimalFormat df = new DecimalFormat("0.0");
					   							if (!showIncompleted || percComplete < 100) {
					   								if (lastActivity.indexOf(",") > -1) {
					   									String[] lastActivityArr = lastActivity.split(",");
					   									Arrays.sort(lastActivityArr, new org.ce.util.NaturalOrderComparator());
					   									lastActivity = lastActivityArr[lastActivityArr.length-1];
					   								}
					   								lastActivity = lastActivity.replaceAll("'","");

					   								if (lastRegionProgram.equals("") || regionId == null || !lastRegionProgram.equals(regionId+" "+programId)) {
					   									if (!jObj3.isEmpty()) {
					   										double percComplete2 = Math.round((((Integer)jObj3.get("numComplete")).floatValue() / ((Integer)jObj3.get("numUsers")).floatValue() * 100) * 10) / 10.0;
					   										if (percComplete2 < 100) {
					   											inProgress++;
					   										} else {
					   											totalComplete++;
					   										}
					   										rptTotalRecords++;
					   										jObj3.put("percComplete",df.format(percComplete2));
					   									}
	
					   									jObj3 = new JSONObject();
					   									if (regionId != null) {
					   										jObj3.put("regionId",regionId);
					   										jObj3.put("region",regionName);
					   										jObj3.put("locationId",null);
					   										jObj3.put("location","");
					   										jObj3.put("programId",programId);
					   										jObj3.put("program",programName);
					   										jObj3.put("numUsers",numUsers.intValue());
					   										jObj3.put("numComplete",numComplete.intValue());
					   										jObj3.put("percComplete",df.format(percComplete));
					   										jObj3.put("assigned",programAssigned);
					   										jObj3.put("deadline",programDeadline);
					   										jObj3.put("lastActivity",lastActivity);
					   										lastRegionProgram = regionId+" "+programId;
					   									} else {
					   										lastRegionProgram = "";
					   									}
					   								} else if (!jObj3.length() ) {
					   									jObj3.put("numUsers",((Integer)jObj3.get("numUsers")) + numUsers.intValue());
					   									jObj3.put("numComplete",((Integer)jObj3.get("numComplete")) + numComplete.intValue());
					   									String lastActivity2 = (String)jObj3.get("lastActivity");
					   									if (lastActivity2.equals("")) {
					   										jObj3.put("lastActivity",lastActivity);
					   									} else if (!lastActivity.equals("")) {
					   										Date lastDate = dateTimeFormat.parse(lastActivity);
					   										Date lastDate2 = dateTimeFormat.parse(lastActivity2);
					   										Calendar lastDateCal = Calendar.getInstance();
					   										Calendar lastDateCal2 = Calendar.getInstance();
					   										lastDateCal.setTime(lastDate);
					   										lastDateCal2.setTime(lastDate2);
					   										if (lastDateCal.after(lastDateCal2)) {
					   											jObj3.put("lastActivity",lastActivity);
					   										}
					   									}
					   								}

					   								if (percComplete < 100) {
					   									inProgress++;
					   								} else {
					   									totalComplete++;
					   								}
					   								rptTotalRecords++;
					   								jObj2 = new JSONObject();
					   								jObj2.put("regionId",regionId);
					   								jObj2.put("region",regionName);
					   								jObj2.put("locationId",locationId);
					   								jObj2.put("location",ucObj[4].toString());
					   								jObj2.put("programId",programId);
					   								jObj2.put("program",ucObj[1].toString());
					   								jObj2.put("numUsers",numUsers.intValue());
					   								jObj2.put("numComplete",numComplete.intValue());
					   								jObj2.put("percComplete",df.format(percComplete));
					   								jObj2.put("assigned",ucObj[9].toString());
					   								jObj2.put("deadline",ucObj[8].toString());
					   								jObj2.put("lastActivity",lastActivity);
					   								jList2.add(jObj2);
					   							}
					   						}
					   					}
					   					if (!jObj3.isEmpty()) {
					   						rptTotalRecords++;
					   						jList1.add(jObj3);
					   					}
					   			}
					   			
					   			
					   			
					   			
		
					   			List list = lookup_mgr.getResults(jObj1.getString("getStatuQuery").replaceAll("%%userQueryString%%", recipient.getQueryString("uc.user.")));
								if (list != null && list.size() > 0){
									for (int a = 0; a < list.size(); a++) {
										Object[] ucObj = (Object[])list.get(a);
										if (ucObj != null) {
											String locName = ucObj[8].toString();
											String username = ucObj[3].toString();
											Long curId = (Long)ucObj[0];
											Long locationId = (Long)ucObj[7];
											int statusId = (Integer)ucObj[2];
											int overallStatus = 1;
							                boolean programComplete = (Boolean)ucObj[13];
											if (statusType.equals("1")) {
												if (programComplete) {
													totalComplete++;
													courseStatus = resource.getString("label.completed");
												} else {
													inProgress++;
													courseStatus = resource.getString("ucsd.status.notcomplete");
												}
												if (!showIncompleted || !programComplete) {
													rptTotalRecords++;
							  						eRow = sheet.createRow(rowNum++);
							  	   					cellNum=0;
							  	   					addCell(ucObj[4].toString()+" "+ucObj[5].toString());
								   		   			if (jObj2.has("showUsernames") && jObj2.getBoolean("showUsernames")) {
														addCell(username);
								   		   			}
													addCell(ucObj[14].toString()+" "+ucObj[15].toString());
													addCell(locName);
													addCell(ucObj[1].toString());
													addCell(courseStatus);
													addCell(ucObj[12]!=null&&ucObj[12].toString().length()>=10?DateUtil.getDate(ucObj[12].toString().substring(0,10), date_format):"");
													addCell(DateUtil.getDate(ucObj[10].toString(), date_format));
												}
											} else {
												totalPrograms++;
												List ucsList = lookup_mgr.getResults(jObj1.getString("getStatuQueryUCS").replaceAll("%%username%%", username).replaceAll("%%curId%%", curId.toString()));
												if (ucsList != null && ucsList.size() > 0) {
													for (int b = 0; b < ucsList.size(); b++) {
														Object[] ucsObj = (Object[])ucsList.get(b);
														if (ucsObj != null) {
															String date = "";
															Long groupId = (Long)ucsObj[1];
															int courseStatusWeight = (Integer)ucsObj[0];
															if (courseStatusWeight < 4 && ucsObj[2] != null) {
																UserCourseResult ucr = ucr_mgr.getUserCourseResult(ucsObj[2].toString());
																if (ucr != null && ucr.getStart_date() != null) {
																	Date startedDate = dateFormat.parse(ucr.getStart_date());
																	Calendar startedDateCal = Calendar.getInstance();
																	startedDateCal.setTime(startedDate);
																	if (startedDateCal.after(startDateCal)) {
																		courseStatusWeight = 4;
																	} else if (ucr.getEnd_date() != null) {
																		Date endDate = dateFormat.parse(ucr.getEnd_date());
																		Calendar endDateCal = Calendar.getInstance();
																		endDateCal.setTime(endDate);
																		if (endDateCal.after(startDateCal)) {
																			courseStatusWeight = 3;
																			date = dateFormat.format(startedDateCal.getTime());
																		} else {
																			date = dateFormat.format(endDateCal.getTime());
																		}
																	} else {
																		courseStatusWeight = 3;
																		date = dateFormat.format(startedDateCal.getTime());
																	}
																}
															}
															if (courseStatusWeight == 1) {
																totalComplete++;
																courseStatus = resource.getString("label.completed");
															} else if (courseStatusWeight == 4) {
																notStarted++;
																courseStatus = resource.getString("ucsd.status.notstarted");
															} else {
																inProgress++;
																courseStatus = resource.getString("label.inprogress");
															}
															if (b == 0) {
																overallStatus = courseStatusWeight;
															} else if (overallStatus < courseStatusWeight || (courseStatusWeight != 4 && overallStatus == 4)) {
																overallStatus = 3;
															}
															rptTotalRecords++;
									  						eRow = sheet.createRow(rowNum++);
									  	   					cellNum=0;
									  	   					addCell(ucObj[4].toString()+" "+ucObj[5].toString());
										   		   			if (jObj2.has("showUsernames") && jObj2.getBoolean("showUsernames")) {
																addCell(username);
										   		   			}
															addCell(ucObj[14].toString()+" "+ucObj[15].toString());
															addCell(locName);
															addCell(ucObj[1].toString());
															addCell(lookup_mgr.getCourseName(carrierId,groupId,locale));
															addCell(courseStatus);
															addCell(ucObj[12]!=null&&ucObj[12].toString().length()>=10?DateUtil.getDate(ucObj[12].toString().substring(0,10), date_format):"");
															addCell(DateUtil.getDate(date, date_format));
															addCell(DateUtil.getDate(ucObj[10].toString(), date_format));
														}
													}
												}
												if (overallStatus == 1) {
													pTotalComplete++;
												} else if (overallStatus == 4) {
													pNotStarted++;
												} else {
													pInProgress++;
												}
											}
								   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")){
							   		   				addUserInfo(username,ucObj[16].toString());
								   			}
										}
									}
								}
								if (includeIndividualModules) {
									List ucsList2 = lookup_mgr.getResults(jObj1.getString("getStatuQueryIndividualModules").replaceAll("%%userQueryString%%", recipient.getQueryString("ucs.user.")));
									if (ucsList2 != null && ucsList2.size() > 0) {
										for (int c = 0; c < ucsList2.size(); c++) {
											Object[] ucsObj = (Object[])ucsList2.get(c);
											if (ucsObj != null) {
												String date = "";
												String locName = ucsObj[8].toString();
												String username = ucsObj[3].toString();
												Long locationId = (Long)ucsObj[7];
							
												Subcompany subcompany = lookup_mgr.getSubcompany(locationId);
												if (subcompany != null && subcompany.getName() != null && !subcompany.getName().equals("")) {
													locName = subcompany.getName()+" - "+locName;
												}
							
												Long groupId = (Long)ucsObj[1];
												int courseStatusWeight = (Integer)ucsObj[0];
												if (courseStatusWeight < 4 && ucsObj[2] != null) {
													UserCourseResult ucr = ucr_mgr.getUserCourseResult(ucsObj[2].toString());
													if (ucr != null && ucr.getStart_date() != null) {
														Calendar startedDate = DateUtil.convertStringToCalendar(ucr.getStart_date());
														if (startedDate.after(startDateCal)) {
															courseStatusWeight = 4;
														} else if (ucr.getEnd_date() != null) {
															Calendar endDate = DateUtil.convertStringToCalendar(ucr.getEnd_date());
															if (endDate.after(startDateCal)) {
																courseStatusWeight = 3;
																date = dateFormat.format(startedDate.getTime());
															} else {
																date = dateFormat.format(endDate.getTime());
															}
														} else {
															courseStatusWeight = 3;
															date = dateFormat.format(startedDate.getTime());
														}
													}
												}
												if (courseStatusWeight == 1) {
													totalComplete++;
													courseStatus = resource.getString("label.completed");
												} else if (courseStatusWeight == 4) {
													notStarted++;
													courseStatus = resource.getString("ucsd.status.notstarted");
												} else {
													inProgress++;
													courseStatus = resource.getString("label.inprogress");
												}
												rptTotalRecords++;
												eRow = sheet.createRow(rowNum++);
						  	   					cellNum=0;
						  	   					addCell(ucsObj[4].toString()+" "+ucsObj[5].toString());
							   		   			if (jObj2.has("showUsernames") && jObj2.getBoolean("showUsernames")) {
													addCell(username);
							   		   			}
												addCell("");
												addCell(locName);
												addCell("");
												addCell(lookup_mgr.getCourseName(carrierId,groupId,locale));
												addCell(courseStatus);
												addCell(ucsObj[9]!=null&&ucsObj[9].toString().length()>=10?ucsObj[9].toString().substring(0,10):"");
												addCell(date);
									   			if (jObj2.has("showPersonalInfo") && jObj2.getBoolean("showPersonalInfo")){
							   		   				addUserInfo(username,ucsObj[10].toString());
								   			}

										}
									}
								}
							}
							if (statusType.equals("2")) {
								notes.append(resource.getString("statusReport.legend.title1") + " : " + 
									resource.getString("label.complete")+": "+df2.format(((float)pTotalComplete/totalPrograms*10000)/100f)+"%; "+
									resource.getString("label.inprogress")+": "+df2.format(((float)pInProgress/totalPrograms*10000)/100f)+"%; "+
									resource.getString("ucsd.status.notstarted")+": "+df2.format(((float)pNotStarted/totalPrograms*10000)/100f)+"%<br/>");
								notes.append(resource.getString("statusReport.legend.title2") + " : " + 
									resource.getString("label.complete")+": "+df2.format(((float)totalComplete/rptTotalRecords*10000)/100f)+"%; "+
									resource.getString("label.inprogress")+": "+df2.format(((float)inProgress/rptTotalRecords*10000)/100f)+"%; "+
									resource.getString("ucsd.status.notstarted")+": "+df2.format(((float)notStarted/rptTotalRecords*10000)/100f)+"%");
							}
				   			
		  	   				}else if (reportType==6){	


		  	   					
		  	   					int rptTotalRecords = 0, totalPrograms = 0, timezoneOffset = 0, totalComplete = 0, inProgress = 0, notStarted = 0, pTotalComplete = 0, pInProgress = 0, pNotStarted = 0;
		  	   					String courseStatus = "";
		  	   				
			  	   				NumberFormat df = NumberFormat.getNumberInstance();
				  	   			df.setMaximumFractionDigits(0);
				  	   			df.setMinimumFractionDigits(0);
				  	   			df.setGroupingUsed(false);
		
		  	   		   			
		   		   				if (allCourse){
			   		   				params.add(allCourseStr);
		   		   				}else{
		   		   					params.addAll(courseParams);
		   		   				}
		   		   				if (allUser){
			   		   				params.add(allUserStr);
		   		   				}else{
		   		   					params.addAll(userParams);
		   		   				}

		   		   				List dates = getDates(jObj2);
		   		   				params.add(getDateFilter(dates,locale,date_format));
		   		   				params.addAll(otherParams);
   		   				

		   					List list = lookup_mgr.getResults(jObj1.getString("getSurveyQuery"));
		   					if (list != null && list.size() > 0) {
		   						String queryString2 = jObj1.getString("getSurveyQuery2").replaceAll("%%questions%%", StringUtils.join(list.toArray(),",")).replaceAll("%%from%%", (String)dates.get(0)).replaceAll("%%to%%", (String)dates.get(1));
		
		   						Map auaCount = new HashMap();
		   						Map qaCount = new HashMap();
		   						Map feedbackCount = new HashMap();
		   						boolean showResults = false;
		   						Long prevAId = null;
		   						Long prevQId = null;
		   						Long qTotal = new Long(0);
		   						List list2 = lookup_mgr.getResults("select aua.questionId, aua.user_response"+queryString2+" group by aua.user_response order by aua.questionId");
		   						for (int a = 0; a < list2.size(); a++) {
		   							Object[] oauaCount = (Object[])list2.get(a);
		   							Long curQId = (Long)oauaCount[0];
		   							Long count = (Long)oauaCount[2];
		   							if (qTotal > 1) {
		   								showResults = true;
		   							}
		   							if (a > 0 && prevQId.longValue() != curQId.longValue()) {
		   								qaCount.put(prevQId,qTotal);
		   								qTotal = new Long(0);
		   							}
		   							prevQId = curQId;
		   							qTotal += count;
		   							auaCount.put((String)oauaCount[1],count);
		   							if (a+1 == list2.size()) {
		   								qaCount.put((Long)prevQId,qTotal);
		   							}
		   						}
		   						if (qTotal > 1) {
		   							showResults = true;
		   						}
		
		   						if (!showResults) {
		   							list2 = lookup_mgr.getResults(jObj1.getString("getSurveyCCA"));
		   							if (list2 != null && list2.size() > 0) {
		   								showResults = true;
		   							}
		   						}
		
		   						if (showResults) {
		   							list2 = lookup_mgr.getResults("select aua.questionId"+queryString2+" and aua.feedback is not null and aua.feedback!='' group by aua.questionId");
		   							for (int a = 0; a < list2.size(); a++) {
		   								Object[] ofeedbackCount = (Object[])list2.get(a);
		   								feedbackCount.put(ofeedbackCount[0],(Long)ofeedbackCount[1]);
		   							}
		
		   							int qNum = 0;
		   							for (int c = 0; c < list.size(); c++) {
		   								Question question = question_mgr.getQuestion(list.get(c).toString());
		   								if (question != null) {
		   									int qType = question.getQuestion_type() != null ? question.getQuestion_type().intValue() : 0;
		   									int qaaCount = 0;
		   									Long complete = (Long)qaCount.get(question.getId());
		   									Long responses = new Long(0);
		   									boolean hasRows = false;
		   									boolean isRow = false;
		   									complete = complete != null ? complete : new Long(0);
		   									if (c == 0 || question.getAssessmentId() != prevAId.longValue()) {
		   										if (c > 0) {
		   											qNum = 0;
		   	   				  						eRow = sheet.createRow(rowNum++);
		   										}
		
		   				  						eRow = sheet.createRow(rowNum++);
		   				  	   					cellNum=0;
		   										Assessment assessment = question.getAssessment();
			   									addCell(" ");
			   									addCell(lookup_mgr.getCourseName(carrierId,assessment.getCourse()));
			   			   		   				String[][] reportTitleExcel = getReportTitle(reportType,"",locale);
			   					  				eRow = sheet.createRow(rowNum++);
			   					  	   			cellNum=0;
			   				  	   				eRow.setHeightInPoints(titleHeight);
			   		  				   			for (int a=0; a < reportTitleExcel.length; a++){
			   		  								addCell(reportTitleExcel[a][0],"headerCell");
			   				  	   					sheet.setColumnWidth(cellNum-1, Integer.valueOf(reportTitleExcel[a][2]));
			   		  				   			}						
		   									}
		   									prevAId = question.getAssessmentId();
		
		   									if (qType == 4) {
		   										List groupWith = lookup_mgr.getResults("select id from QuestionGroupMember where status=1 and questionId="+question.getId()+" and question_group_id=8");
		   										if (groupWith != null && groupWith.size() > 0) {
		   											hasRows = true;
		   											isRow = true;
		   										} else if ((c+1) < list.size()){
		   											Question nextQuestion = question_mgr.getQuestion(list.get(c+1).toString());
		   											if (nextQuestion != null) {
		   												groupWith = lookup_mgr.getResults("select id from QuestionGroupMember where status=1 and questionId="+nextQuestion.getId()+" and question_group_id=8");
		   												if (groupWith != null && groupWith.size() > 0) {
		   													hasRows = true;
		   												}
		   											}
		   										}
		   									}
		   									
		
		   									Long fCount = (Long)feedbackCount.get(question.getId());
		
		   									qNum++;
					  						eRow = sheet.createRow(rowNum++);
					  	   					cellNum=0;
											addCell(String.valueOf(qNum));
											addCell(Jsoup.parse(StringEscapeUtils.unescapeHtml3(question.getQuestion_text())).text());
											
		   									list2 = lookup_mgr.getResults("select id, (CASE WHEN answer_text IS NULL THEN '' ELSE answer_text END) from Answer where questionId="+question.getId()+" order by correct");
		   									if (list2 != null && list2.size() > 0) {
		   										for (int d = 0; d < list2.size(); d++) {
		   											Object[] aObj = (Object[])list2.get(d);
		   											if (aObj != null) {
		   												Long aId = (Long)aObj[0];
		   												double percent = 0.0;
		   												double avg = 0.0;
		   												Long count = (Long)auaCount.get(aId.toString());
		   												count = count != null ? count : 0;
		
		   												if (!isRow) {
		   													responses = count;
		   													if (complete > 0 && responses > 0) {
		   														percent = Math.round((responses.floatValue() / complete * 100) * 10) / 10.0;
		   													}
		   												} else {
		   													responses += count;
		   													qaaCount += (d+1)*count;
		   													if (list2.size() == d+1) {
		   														if (responses > 0) {
		   															avg = Math.round((qaaCount / responses) * 10) / 10.0;
		   															percent = Math.round((avg / (d+1) * 100) * 10) / 10.0;
		   														}
		   													}
		   												}
		   												if (!isRow || list2.size() == d+1) {
		   							  						eRow = sheet.createRow(rowNum++);
		   							  	   					cellNum=0;
		   													addCell(" ");
		   													addCell(Jsoup.parse(StringEscapeUtils.unescapeHtml3(aObj[1].toString())).text());
		   													addCell(df.format(responses));
		   													addCell(df.format(percent)+"%");
		   												}
		   											}
		   										}
		   									}
		
		   								}
		   								rptTotalRecords++;
		   							}
		   						}
		   					}
		
		   				}
		   				
		   				if (rowNum>1){
				   		   	filename = CEProperties.getInstance().getProperty("webroot.dir")+"/files/"+StringUtil.trimFilename(savedReport.getName().replaceAll("/", "_").replaceAll("\\\\", "_").replaceAll("<", "_").replaceAll(">", "_")+".xlsx");
				   		   	emailTitle = resource.getString("company.name")+" "+resource.getString("report.scheduled")+": "+savedReport.getName();
				   		   	emailTemplate = "/org/ce/scheduledtask/ScheduledReport"+(locale.toString().toLowerCase().equals("fr")?"_fr":"")+".vm";
				   		   	os = new FileOutputStream(filename);
				   		   	wb.write(os);
		   				
					      	model.put("title", emailTitle);
					      	model.put("filters", params);
					      	model.put("notes", notes.toString());
		   				   	model.put("reportName", savedReport.getName());
		   				   	model.put("date", DateUtil.getDate(curr_date, date_format));
		   				}
   	   				}

   	   				if ((roleLevel!=null && roleLevel.intValue()>=40 || rowNum>1) && new File(filename).exists()){  
	   	   				model.put("firstName", recipient.getFirstName());
						sendvelocityMail.sendvelocityMail(recipient, emailTitle, emailTemplate, model,filename, filename.substring(filename.lastIndexOf("/")+1) );
						msg_log.append(emailTitle+" sent to "+recipient.getFullName()+"&nbsp;&lt;"+recipient.getEmails()+"&gt; <br />" );
   	   				}
   	   			}

   	   			scheduledReport.setModified(curr_date);
				scheduledReport.setModified_by("batchadmin");
				scheduledReport.setNext_report(getNextReport(scheduledReport.getNext_report(), scheduledReport.getFrequency()));
				scheduledReport_mgr.saveScheduledReport(scheduledReport);
				
				if (new File(filename).exists()){
					new File(filename).delete();
				}
				
   			}
   		} catch(Exception e) {
   			e.printStackTrace();
   		} finally {
   			if (os != null) {
   				try {
   					os.close();
   				} catch (IOException e) {}
   			}
   			if (wb != null) {
   				try {
   					wb.close();
   					wb.dispose();
   				} catch (IOException e) {}
   			}
   		}
	   	
		model = new HashMap();
		model.put("title", "Scheduled Reports");
		model.put("date", "");
		model.put("msg", msg_log.toString());
	   	if (msg_log!=null && msg_log.length()>10){
			SendvelocityMail sendMail=new SendvelocityMail();	
   			User sys_account=(User)user_mgr.getUser("julie.yuan");
   			if ( sys_account.canNotified()){
				sendMail.sendvelocityMail(sys_account,"Scheduled Reports" ,"/org/ce/scheduledtask/OTABilling.vm" ,model);
   			}
   		}
   		

		System.out.println("---------- Scheduled Reports End ---------- " + new Date());

	}	

}

	

