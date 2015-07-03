<!--#include file="../../easyasp/easp.asp" --><%
Dim datetime
datetime = "2013-12-22 23:34:45"
Easp.print("给定日期是：")
Easp.println Easp.Date.Format(datetime, "y-mm-dd hh:ii:ss 星期w")
Easp.print("从现在算起是：")
Easp.println Easp.Date.Format(datetime, Now)
Easp.print("这个月的第一天是：")
Easp.println Easp.Date.FirstDayOfMonth(datetime)
Easp.print("这个月的最后一天是：")
Easp.println Easp.Date.LastDayOfMonth(datetime)
Easp.print("如果每周从星期一开始，这周的第一天是：")
Easp.println Easp.Date.FirstDayOfWeek(datetime)
'从星期日开始一周
Easp.Date.WeekStarting = 1
Easp.print("如果每周从星期日开始，这周的最后一天是：")
Easp.println Easp.Date.LastDayOfWeek(datetime)
Easp.print("这个时间转成Unix时间戳是：")
Easp.println Easp.Date.ToUnixTimeCn(datetime)
%>