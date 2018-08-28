<%@include file="../helpers/reportLookups.jsp"%>
<%@page import="java.net.URLEncoder"%>
<%@page import="java.text.DecimalFormat"%>
<%@page import="java.text.NumberFormat"%>
<%@page import="java.text.SimpleDateFormat"%>
<%@page import="java.util.Calendar"%>
<%@page import="java.util.Date"%>
<%@page import="javax.servlet.jsp.jstl.core.Config"%>
<%@page import="org.ce.model.Subcompany"%>
<%@page import="org.ce.model.UserCourseStatus"%>
<%@page import="org.ce.model.UserCourseStatusRank"%>
<%@page import="org.ce.model.UserInfo"%>
<%@page import="org.ce.service.LookupManager"%>
<%@page import="org.ce.service.UserCourseResultManager"%>
<%@page import="org.ce.service.UserInfoManager"%>
<%@page import="org.ce.util.DateUtil"%>
<%@page import="org.ce.webapp.action.UTF8Control"%>
<%@page import="org.json.simple.JSONArray"%>
<%@page import="org.json.simple.JSONObject"%>
<%@page import="org.springframework.context.ApplicationContext"%>
<%@page import="org.springframework.web.context.support.WebApplicationContextUtils"%>
<%
User current_user = (User)session.getAttribute("currentUserForm");
ApplicationContext ctx = WebApplicationContextUtils.getRequiredWebApplicationContext(getServletContext());

LookupManager lookup_mgr = (LookupManager) ctx.getBean("lookupManager");
UserCourseResultManager ucr_mgr = (UserCourseResultManager) ctx.getBean("userCourseResultManager");
UserInfoManager userInfo_mgr = (UserInfoManager) ctx.getBean("userInfoManager");

Locale locale = (Locale)Config.get(session, Config.FMT_LOCALE);
ResourceBundle resource = lookup_mgr.getResourceBundle(locale);

boolean showInactive = false;
boolean showIncompleted = false;
boolean includeIndividualModules = false;
boolean showUnassignedPrograms = false;
boolean carriersAll = false;
boolean summarizeByPlace = false;
String queryString = "";
String calStart = "";
String todayDate = "";
String statusType = "1";
String courseStatus = "";
String lastRegionProgram = "";
String[] userList = null;
String[] courseList = null;
int rptTotalRecords = 0;
int totalPrograms = 0;
int timezoneOffset = 0;
int totalComplete = 0;
int inProgress = 0;
int notStarted = 0;
int pTotalComplete = 0;
int pInProgress = 0;
int pNotStarted = 0;
ReportLookup rl = new ReportLookup(locale);
JSONObject jObj1 = new JSONObject();
JSONObject jObj2 = new JSONObject();
JSONObject jObj3 = new JSONObject();
JSONArray jList1 = new JSONArray();
JSONArray jList2 = new JSONArray();
List list = null;

Calendar now = Calendar.getInstance();
SimpleDateFormat generated_time = new SimpleDateFormat("yyyy-MM-dd hh:mm aa");
SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
SimpleDateFormat dateTimeFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

if (request.getParameter("statusType") != null && !request.getParameter("statusType").equals("")){
	statusType = request.getParameter("statusType");
}

if (request.getParameter("userProgress") != null && !request.getParameter("userProgress").equals("")){
	userList = request.getParameter("userProgress").split(",");
}

if (request.getParameter("programsProgress") != null && !request.getParameter("programsProgress").equals("")){
	courseList = request.getParameter("programsProgress").split(",");
}

if (request.getParameter("showInactive") != null && request.getParameter("showInactive").equals("true")){
	showInactive = true;
}

if (request.getParameter("showIncompleted") != null && request.getParameter("showIncompleted").equals("true")){
	showIncompleted = true;
}

if (statusType.equals("2") && request.getParameter("includeIndividualModules") != null && request.getParameter("includeIndividualModules").equals("true")){
	includeIndividualModules = true;
}

if (request.getParameter("showUnassignedPrograms") != null && request.getParameter("showUnassignedPrograms").equals("true")){
	showUnassignedPrograms = true;
}

if (statusType.equals("1") && request.getParameter("summarizeByPlace") != null && request.getParameter("summarizeByPlace").equals("true")){
	summarizeByPlace = true;
}

if (request.getParameter("calStart") != null && !request.getParameter("calStart").equals("")){
	calStart = request.getParameter("calStart");
} else {
	calStart = dateFormat.format(now.getTime());
}

if (request.getParameter("timezoneOffset") != null && !request.getParameter("timezoneOffset").equals("")){
	timezoneOffset = Integer.parseInt(request.getParameter("timezoneOffset"));
}

todayDate = dateFormat.format(now.getTime());
Date startDate = dateFormat.parse(calStart);
Calendar startDateCal = Calendar.getInstance();
startDateCal.setTime(startDate);
startDateCal.set(Calendar.HOUR_OF_DAY, 23);
startDateCal.set(Calendar.MINUTE, 59);
startDateCal.set(Calendar.SECOND, 59);
startDateCal.set(Calendar.MILLISECOND, 999);
if (!todayDate.equals(calStart)) {
	startDateCal.add(Calendar.MILLISECOND, timezoneOffset);
}
calStart = dateTimeFormat.format(startDateCal.getTime());

response.setContentType("application/json");
response.setHeader("Cache-Control", "no-cache");

if (current_user != null && current_user.getRoleLevel() > 0) {
	if (summarizeByPlace) {
		list = lookup_mgr.getResults(rl.getStatuQueryPlace(current_user, userList, courseList, showUnassignedPrograms, showInactive, showIncompleted, carriersAll, current_user.getQueryString("uc.user.")));
		if (list != null && list.size() > 0){
			for (int a = 0; a < list.size(); a++) {
				Object[] ucObj = (Object[])list.get(a);
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
								jList1.add(jObj3);
							}
							jList1.addAll(jList2);
							jList2 = new JSONArray();
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
						} else if (!jObj3.isEmpty()) {
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
			jList1.addAll(jList2);
		}
	} else {
		list = lookup_mgr.getResults(rl.getStatuQuery(current_user, userList, courseList, showUnassignedPrograms, showInactive, showIncompleted, carriersAll ,current_user.getQueryString("uc.user.")));
		if (list != null && list.size() > 0){
			for (int a = 0; a < list.size(); a++) {
				Object[] ucObj = (Object[])list.get(a);
				if (ucObj != null) {
					String locName = ucObj[8].toString();
					String username = ucObj[3].toString();
					String phone = ucObj[16].toString();
					Long curId = (Long)ucObj[0];
					Long locationId = (Long)ucObj[7];
					Long carrierId = (Long)ucObj[9];
					int statusId = (Integer)ucObj[2];
					int overallStatus = 1;
					boolean programComplete = (Boolean)ucObj[13];
					boolean canSendEmail = false;
					if (!programComplete) {
						canSendEmail = !ucObj[11].toString().equals("");
						if (!canSendEmail) {
							List eNotifications = lookup_mgr.getResults("select count(*) from NotificationAddress where username='"+username+"'");
							if (eNotifications != null && ((Long)eNotifications.get(0)).intValue() > 0) {
								canSendEmail = true;
							}
						}
					}
					if (statusType.equals("1")) {
						if (!showIncompleted || !programComplete) {
							if (programComplete) {
								totalComplete++;
								courseStatus = resource.getString("label.completed");
							} else {
								inProgress++;
								courseStatus = resource.getString("ucsd.status.notcomplete");
							}
							rptTotalRecords++;
							jObj2 = new JSONObject();
							jObj2.put("username",username);
							jObj2.put("firstName",ucObj[4].toString());
							jObj2.put("lastName",ucObj[5].toString());
							jObj2.put("m_firstName",ucObj[14].toString());
							jObj2.put("m_lastName",ucObj[15].toString());
							jObj2.put("status",(Integer)ucObj[6]);
							jObj2.put("locationId",locationId);
							jObj2.put("location",locName);
							jObj2.put("programId",curId);
							jObj2.put("program",ucObj[1].toString());
							jObj2.put("programComplete",programComplete);
							jObj2.put("courseStatus",courseStatus);
							jObj2.put("assigned",ucObj[12].toString());
							jObj2.put("deadline",ucObj[10].toString());
							jObj2.put("canSendEmail",canSendEmail);
							UserInfo userInfo = userInfo_mgr.getUserInfoByUsername(username);
							if (userInfo != null) {
								jObj2.put("address",(userInfo.getAddress() != null ? userInfo.getAddress() : ""));
								jObj2.put("city",(userInfo.getCity() != null ? userInfo.getCity() : ""));
								jObj2.put("province",(userInfo.getProvince() != null ? userInfo.getProvince() : ""));
								jObj2.put("postalCode",(userInfo.getPostal() != null ? userInfo.getPostal() : ""));
								jObj2.put("license",(userInfo.getLicensenumber() != null ? userInfo.getLicensenumber() : ""));
								jObj2.put("licenseProv",(userInfo.getLicense_prov() != null ? userInfo.getLicense_prov() : ""));
								jObj2.put("jobTitle",(userInfo.getJob_title() != null ? userInfo.getJob_title() : ""));
								jObj2.put("phone",!phone.equals("") ? phone : (userInfo.getPhone_number() != null ? userInfo.getPhone_number() : ""));
							}
							jList1.add(jObj2);
						}
					} else {
						totalPrograms++;
						List ucsList = lookup_mgr.getResults(rl.getStatuQueryUCS(current_user, username, curId.toString(), showIncompleted, carriersAll));
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
											Date startedDate = dateTimeFormat.parse(ucr.getStart_date());
											Calendar startedDateCal = Calendar.getInstance();
											startedDateCal.setTime(startedDate);
											if (startedDateCal.after(startDateCal)) {
												courseStatusWeight = 4;
											} else if (ucr.getEnd_date() != null) {
												Date endDate = dateTimeFormat.parse(ucr.getEnd_date());
												Calendar endDateCal = Calendar.getInstance();
												endDateCal.setTime(endDate);
												if (endDateCal.after(startDateCal)) {
													courseStatusWeight = 3;
													date = dateTimeFormat.format(startedDateCal.getTime());
												} else {
													date = dateTimeFormat.format(endDateCal.getTime());
												}
											} else {
												courseStatusWeight = 3;
												date = dateTimeFormat.format(startedDateCal.getTime());
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
									jObj2 = new JSONObject();
									jObj2.put("username",username);
									jObj2.put("firstName",ucObj[4].toString());
									jObj2.put("lastName",ucObj[5].toString());
									jObj2.put("m_firstName",ucObj[14].toString());
									jObj2.put("m_lastName",ucObj[15].toString());
									jObj2.put("status",(Integer)ucObj[6]);
									jObj2.put("locationId",locationId);
									jObj2.put("location",locName);
									jObj2.put("programId",curId);
									jObj2.put("program",ucObj[1].toString());
									jObj2.put("programComplete",programComplete);
									jObj2.put("moduleId",groupId);
									jObj2.put("module",lookup_mgr.getCourseName1(current_user,groupId));
									jObj2.put("courseStatusId",courseStatusWeight);
									jObj2.put("courseStatus",courseStatus);
									jObj2.put("date",date);
									jObj2.put("assigned",ucObj[12].toString());
									jObj2.put("deadline",ucObj[10].toString());
									jObj2.put("canSendEmail",canSendEmail);
									UserInfo userInfo = userInfo_mgr.getUserInfoByUsername(username);
									if (userInfo != null) {
										jObj2.put("address",(userInfo.getAddress() != null ? userInfo.getAddress() : ""));
										jObj2.put("city",(userInfo.getCity() != null ? userInfo.getCity() : ""));
										jObj2.put("province",(userInfo.getProvince() != null ? userInfo.getProvince() : ""));
										jObj2.put("postalCode",(userInfo.getPostal() != null ? userInfo.getPostal() : ""));
										jObj2.put("license",(userInfo.getLicensenumber() != null ? userInfo.getLicensenumber() : ""));
										jObj2.put("licenseProv",(userInfo.getLicense_prov() != null ? userInfo.getLicense_prov() : ""));
										jObj2.put("jobTitle",(userInfo.getJob_title() != null ? userInfo.getJob_title() : ""));
										jObj2.put("phone",!phone.equals("") ? phone : (userInfo.getPhone_number() != null ? userInfo.getPhone_number() : ""));
									}
									jList1.add(jObj2);
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
				}
			}
		}
	}
	if (includeIndividualModules) {
		List ucsList2 = lookup_mgr.getResults(rl.getStatuQueryIndividualModules(current_user, userList, showInactive, showIncompleted, carriersAll, current_user.getQueryString("ucs.user.")));
		if (ucsList2 != null && ucsList2.size() > 0) {
			for (int c = 0; c < ucsList2.size(); c++) {
				Object[] ucsObj = (Object[])ucsList2.get(c);
				if (ucsObj != null) {
					String date = "";
					String phone = ucsObj[12].toString();
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
					jObj2 = new JSONObject();
					jObj2.put("username",username);
					jObj2.put("firstName",ucsObj[4].toString());
					jObj2.put("lastName",ucsObj[5].toString());
					jObj2.put("m_firstName",ucsObj[10].toString());
					jObj2.put("m_lastName",ucsObj[11].toString());
					jObj2.put("status",(Integer)ucsObj[6]);
					jObj2.put("locationId",locationId);
					jObj2.put("location",locName);
					jObj2.put("moduleId",groupId);
					jObj2.put("module",lookup_mgr.getCourseName1(current_user,groupId));
					jObj2.put("courseStatusId",courseStatusWeight);
					jObj2.put("courseStatus",courseStatus);
					jObj2.put("assigned",ucsObj[9].toString());
					jObj2.put("date",date);
					UserInfo userInfo = userInfo_mgr.getUserInfoByUsername(username);
					if (userInfo != null) {
						jObj2.put("address",(userInfo.getAddress() != null ? userInfo.getAddress() : ""));
						jObj2.put("city",(userInfo.getCity() != null ? userInfo.getCity() : ""));
						jObj2.put("province",(userInfo.getProvince() != null ? userInfo.getProvince() : ""));
						jObj2.put("postalCode",(userInfo.getPostal() != null ? userInfo.getPostal() : ""));
						jObj2.put("license",(userInfo.getLicensenumber() != null ? userInfo.getLicensenumber() : ""));
						jObj2.put("licenseProv",(userInfo.getLicense_prov() != null ? userInfo.getLicense_prov() : ""));
						jObj2.put("jobTitle",(userInfo.getJob_title() != null ? userInfo.getJob_title() : ""));
						jObj2.put("phone",!phone.equals("") ? phone : (userInfo.getPhone_number() != null ? userInfo.getPhone_number() : ""));
					}
					jList1.add(jObj2);
				}
			}
		}
	}
	jObj2 = new JSONObject();
	jObj2.put("totalComplete",pTotalComplete);
	jObj2.put("inProgress",pInProgress);
	jObj2.put("notStarted",pNotStarted);
	jObj2.put("totalPrograms",totalPrograms);
	jObj1.put("programCount",jObj2);
	jObj1.put("totalComplete",totalComplete);
	jObj1.put("inProgress",inProgress);
	jObj1.put("notStarted",notStarted);
}
jObj1.put("records",jList1);

Calendar rptDone = Calendar.getInstance();
long rptTime_Long = rptDone.getTimeInMillis() - now.getTimeInMillis();
float rptTime_float = (float) rptTime_Long/1000;
now.add(Calendar.MILLISECOND, -(timezoneOffset));

jObj1.put("totalRecords",rptTotalRecords);
jObj1.put("reportTime",rptTime_float);
jObj1.put("reportTimeNow",generated_time.format(now.getTime()));
jObj1.put("fullName",current_user.getFullName());
jObj1.put("summarizeByPlace",summarizeByPlace);

jObj2 = new JSONObject();
jObj2.put("getStatuQuery", rl.getStatuQuery(current_user, userList, courseList, showUnassignedPrograms, showInactive, showIncompleted, carriersAll, "%%userQueryString%%"));
jObj2.put("getStatuQueryUCS", rl.getStatuQueryUCS(current_user, "%%username%%", "%%curId%%", showIncompleted, carriersAll));
jObj2.put("getStatuQueryIndividualModules", rl.getStatuQueryIndividualModules(current_user, userList, showInactive, showIncompleted, carriersAll, "%%userQueryString%%"));
jObj1.put("savedReportData",jObj2);

response.getWriter().write(jObj1.toJSONString());
%>