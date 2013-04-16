#! /usr/bin/python
# -*- coding: utf-8 -*-

import sys
import MySQLdb
import time
import datetime
import calendar
import os
import os.path

from tempfile import TemporaryFile
from xlwt import Workbook,easyxf
from xlrd import open_workbook

#当天日期
today = datetime.date.today()

#生成报表的目录
report_dir="/work/opt/zabbix-reports"
#MySQL的相关信息
host='127.0.0.1'
port=3306
user = 'report_user'
password = 'report_passwd'
database = 'zabbix'

#要生成执行的KEY，都存到此元组中
keys = ("cpuload","disk_usage","network_in","network_out")
#除boolen型（如vip)的监控项外，都设置了阀值，出报表区间内最大值操作此值的，显示背景红色
#所有没有单独判断的key, 都要在此字典中定义其阀值
#-------------------------------------------------------------
# 网卡流量阀值：50M/s 
#-------------------------------------------------------------
thre_dic = {"cpuload":15,"disk_usage":85,"network_in":409600}

#-------------------------------------------------------------
#自定义生成报表：
#-------------------------------------------------------------
def custom_report(startTime,endTime):
  
	sheetName =  time.strftime('%m%d_%H%M',startTime) + "_TO_" +time.strftime('%m%d_%H%M',endTime)
	customStart = time.mktime(startTime)
	customEnd = time.mktime(endTime)
	generate_excel(customStart,customEnd,0,sheetName)

#-------------------------------------------------------------
# 按日生成报表：
# 	以执行脚本当前unix timestamp和当天午夜的unix timestamp来抽取报表
#	脚本一定要24点之前运行
#-------------------------------------------------------------
def daily_report():
#	today = datetime.date.today() #获取今天日期
	dayStart = time.mktime(today.timetuple()) #由今天日期，取得凌晨unix timestamp
	dayEnd = time.time() #获取当前系统unix timestamp
	sheetName = time.strftime('%Y%m%d',time.localtime(dayEnd))
    	generate_excel(dayStart,dayEnd,1,sheetName)    
#-------------------------------------------------------------
# 按星期生成报表
#-------------------------------------------------------------

def weekly_report():
	lastMonday = today
#	lastMonday = datetime.date.today()#获取今天日期
	#取得上一个星期一
	while lastMonday.weekday() != calendar.MONDAY:
		lastMonday -= datetime.date.resolution
	
	weekStart = time.mktime(lastMonday.timetuple())# 获取周一午夜的unix timestamp
	weekEnd = time.time()#获取当前系统unix timestamp
	#weekofmonth = (datetime.date.today().day+7-1)/7
	weekofmonth = (today.day+7-1)/7
	sheetName = "weekly_" + time.strftime('%Y%m',time.localtime(weekEnd)) + "_" + str(weekofmonth)
	generate_excel(weekStart,weekEnd,2,sheetName)
			
#-------------------------------------------------------------
# 按月生成报表
#-------------------------------------------------------------

def monthly_repport():
#	firstDay =  datetime.date.today() #当前第一天的日期
	firstDay =  today #当前第一天的日期
	#取得当月第一天的日期
	while firstDay.day != 1:
		firstDay -= datetime.date.resolution
	monthStart = time.mktime(firstDay.timetuple()) #当月第一天的unix timestamp
	monthEnd = time.time()	#当前时间的unix timestamp
	sheetName = "monthly_" + time.strftime('%Y%m',time.localtime(monthEnd))
	generate_excel(monthStart,monthEnd,3,sheetName)
	

#-------------------------------------------------------------
#  获取MySQL Connection
#-------------------------------------------------------------
def getConnection():
       # print "准备连接MySQL "
        try:
                connection=MySQLdb.connect(host=host,port=port,user=user,passwd=password,db=database,connect_timeout=1);
        except MySQLdb.Error, e:
                print "Error %d: %s" % (e.args[0], e.args[1])
                sys.exit(1)
	return connection

#-------------------------------------------------------------
# 返回所有主机IP和hostid, 如：('192.168.10.62', 10113L,0),其中Role为添加的字段，1：M, 2:S,3:N
#-------------------------------------------------------------
def getHosts():
	conn=getConnection()
	cursor = conn.cursor()
	command = cursor.execute("""select ip,hostid,Role from hosts where ip<>'127.0.0.1' and ip<>'' and status=0 order by ip;""");
	hosts = cursor.fetchall()
	cursor.close()
	conn.close()
	return hosts

#-------------------------------------------------------------
# 返回指定主机监控Item的itmeid,
#-------------------------------------------------------------
def getItemid(hostid):
	keys_str = "','".join(keys)
	conn=getConnection()
	cursor = conn.cursor()
	command = cursor.execute("""select itemid from items where hostid=%s and key_ in ('%s')""" %(hostid,keys_str));
	itemids =  cursor.fetchall()
	cursor.close()
	conn.close()
	return itemids
#-------------------------------------------------------------
# 返回无指定hostid主机的报表值， 只针对数字history表中
#-------------------------------------------------------------

def getReportById_1(hostid,start,end):
	keys_str = "','".join(keys)
        conn=getConnection()
        cursor = conn.cursor()
	command = cursor.execute("""select items.itemid , key_ as key_value ,units, max(history.value) as max,avg(history.value) as average ,min(history.value) as min  from history, items where items.hostid=%s and items.key_ in ('%s')and items.value_type=0  and history.itemid=items.itemid  and (clock>%s and clock<%s)  group by itemid, key_value;""" %(hostid,keys_str,start,end));	
	values =  cursor.fetchall()
        cursor.close()
	conn.close();
	return values

#-------------------------------------------------------------
# 返回无指定hostid主机的报表值， 只针无符号数history_uint表, items.value_type=3
#-------------------------------------------------------------

def getReportById_2(hostid,start,end):
        keys_str = "','".join(keys)
        conn=getConnection()
        cursor = conn.cursor()
        command = cursor.execute("""select items.itemid , key_ as key_value ,units, max(history_uint.value) as max,avg(history_uint.value) as average ,min(history_uint.value) as min  from history_uint, items where items.hostid=%s and items.key_ in ('%s')and items.value_type=3  and history_uint.itemid=items.itemid and (clock>%s and clock<%s) group by itemid, key_value;""" %(hostid,keys_str,start,end));
        values =  cursor.fetchall()
        cursor.close()
        conn.close();
        return values
#--------------------------------------------------------------
#文件：生成Excel报表 
#参数， start:抽取数据开始时间点 ， end:抽到数据结束时间点
#	reportType:生成报表类型： 1 daily , 2 weekly, 3 monthly
#-----------------------------------------------------------------

def generate_excel(start,end,reportType,sheetName):
	book = Workbook(encoding='utf-8')
	sheet1 = book.add_sheet(sheetName)	
	merge_col = 1
	merge_col_step = 2

	title_col = 1
	title_col_step = 2
	
	hosts = getHosts()
	isFirstLoop=1
	host_row = 2 #host ip所在的行号
	
	max_col = 1
	avg_col = 2
	
	#这义Excel的各种格式
	normal_style = easyxf(
'borders: right thin,top thin,left thin, bottom thin;'
'align: vertical center, horizontal center;'
)
	abnormal_style = easyxf(
'borders: right thin, bottom thin,top thin,left thin;'
'pattern: pattern solid, fore_colour red;'
'align: vertical center, horizontal center;'
)


	sheet1.write_merge(0,1,0,0,"HOSTS")
	for ip,hostid,role in hosts:
		sheet1.row(host_row).set_style(normal_style)
		max_col = 1
	        avg_col = 2
		reports = getReportById_1(hostid,start,end) + getReportById_2(hostid,start,end)
		if(isFirstLoop==1):#第一次循环时，写表头
			sheet1.write(host_row,0,ip,normal_style)
			for report in reports:
				title = report[1] + " " + report[2]		
				sheet1.write_merge(0,0,merge_col,merge_col+1,title,normal_style)
				merge_col += merge_col_step
					
				sheet1.write(1,title_col,"MAX",normal_style)
				sheet1.write(1,title_col+1,"Average",normal_style)
				title_col += title_col_step
		
				#写数据,判断最大值是否超过指定的阀值
				#当最大值大于指定的阀值，此显示为红色
				if(report[3] >= thre_dic[report[1]]):
					sheet1.write(host_row,max_col,report[3],abnormal_style)
					sheet1.write(host_row,avg_col,report[4],normal_style)
				else:	#未超过阀值则正常显示
					sheet1.write(host_row,max_col,report[3],normal_style)
					sheet1.write(host_row,avg_col,report[4],normal_style)
				max_col = max_col + 2
				avg_col =avg_col+ 2
				isFirstLoop=0	
		else:
			sheet1.write(host_row,0,ip,normal_style)
			for report in reports:
				#当最大值大于指定的阀值，此显示为红色
                        	if(report[3] >= thre_dic[report[1]]):
                                	sheet1.write(host_row,max_col,report[3],abnormal_style)
                                        sheet1.write(host_row,avg_col,report[4],normal_style)
                        	 else:   #未超过阀值则正常显示
                                 	sheet1.write(host_row,max_col,report[3],normal_style)
                                 	sheet1.write(host_row,avg_col,report[4],normal_style)
	
                        	max_col = max_col + 2
                                avg_col =avg_col+ 2

		host_row = host_row +1
	saveReport(reportType,book)

#----------------------------------------------------------------------
#函数：根据不同的报表类型，实现不同的保存方式
#参数： reportType 报表类型：0 custom 1 daily, 2 weekly , 3 monthly
#	workBook 当前Excel的工作薄	
#---------------------------------------------------------------------
def saveReport(reportType,workBook):
	#报表目录是否存在，不存在则新创建
	if(not (os.path.exists(report_dir))):
		os.makedirs(report_dir)
	#切换到报表目录
	os.chdir(report_dir)
	#报表每以月单位的目录存放
	month_dir=time.strftime('%Y-%m',time.localtime(time.time()))
	if(not (os.path.exists(month_dir))):
		os.mkdir(month_dir)
	os.chdir(month_dir)
	#自定义生成报表	
	if(reportType == 0):
		excelName = "custom_report_"+ time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) + ".xls"		
	#日报
	elif(reportType == 1):
		excelName = "daily_report_" + time.strftime('%Y%m%d',time.localtime(time.time())) + ".xls"
	#周报
	elif(reportType == 2):
		#weekofmonth = (datetime.date.today().day+7-1)/7			
		weekofmonth = (today.day+7-1)/7			
		excelName = "weekly_report_" +  time.strftime('%Y%m',time.localtime(time.time())) +"_" + str(weekofmonth) + ".xls"
	#月报
	else:
		monthName = time.strftime('%Y%m',time.localtime(time.time()))
		excelName = "monthly_report_" + monthName + ".xls"
#		currentDir = os.getcwd()
#		files = os.listdir(currentDir)#默认为当前目录，也就是月目录
#		for file in files:
#			wb = open_workbook(file)
	print excelName				
	workBook.save(excelName)

#----------------------------------------------------=-----------------
# 入口函数
#------------------------------------------------------------------------

def main():
	#最好加上时间类型检查
	argvCount = len(sys.argv) #参数个数，用于判断是生成自定义报表还是周期性报表	
	dateFormat = "%Y-%m-%d %H:%M:%S"
	today = datetime.date.today()
	if(argvCount == 2):
		#只传入一个参数，生成自定义报表为：当天00点到当前时间的报表
		#时间都格式化为元组格式传入
		startTime = today.timetuple()
		dateFormat = "%Y-%m-%d %H:%M:%S"
		endTime = time.strptime(sys.argv[1],dateFormat) #中止时间为当前时间
		custom_report(startTime,endTime)
	
	elif(argvCount == 3):
		#传入两个参数，生成自定义报表为：以第一参数为起始时间，第二参数为结束时间的区别报表
		startTime =  time.strptime(sys.argv[1],dateFormat)
		endTime =  time.strptime(sys.argv[2],dateFormat)
		custom_report(startTime,endTime)		
	elif(argvCount ==1):
		#无参数传入则生成周期性报表
		today = datetime.date.today()
		dayOfMonth = today.day #取得当天为月的第几天
		#取得年
		year = int(time.strftime('%Y',time.localtime(time.time())))
		#取得月数
		month = int(time.strftime('%m',time.localtime(time.time())))
		#取得当月有多少天
		lastDayOfMonth = calendar.monthrange(year,month)[1]	
		#每天都要生成日报
		daily_report()
		#当前星期天，生成周报
		if(today.weekday()==6): 
			weekly_report()
		#当月最后一天，生成月报
		if(dayOfMonth == lastDayOfMonth):
			monthly_repport()
	else:
		#参数个数大于2为非法情况，打印异常信息，退出报表生成
		usage()
def usage():
	print """脚本没传入参数，则执行周期性报表；参数可为1，2个，注意时间格式强制要求： zabbix-report.py ['2012-09-01 01:12:00'] ['2012-09-01 01:12:00']"""

#运行程序
main()
