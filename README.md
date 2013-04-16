zabbix-report
=============
Description:
  This script for generating zabbix monthly, weekly, daily monitor item report.

Requirements:
    
    1  zabbix 1.8 version and zabbix's MySQL database need to be installed.
    
    2  python2.4 or laster version need to be installed.
    
    3  MySQLdb , xlwt and xlrd need to be installed.

Usage:

    1  Adding the zabbix-report.py script to the zabbix's MySQL database server's crontab.
       e.g   01 00 * * * /work/opt/zabbix-report.py

    2  Updating the MySQL connection variables in the script:
       port=3306
       user = 'report_user'
       password = 'report_passwd'
       database = 'zabbix'
    
    3  Adding the Items into the keys tuple and Item's threshold value in the thre_dic dictionary, for example.
       keys = ("cpuload","disk_usage","network_in","network_out")
       thre_dic = {"cpuload":15,"disk_usage":85,"network_in":409600,"network_out":409600}
