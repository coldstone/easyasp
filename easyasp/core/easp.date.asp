<%
'######################################################################
'## Easp.Date.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Date & Time Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-01-30
'## Description :   Format and processing the date and time object
'##
'######################################################################

Class EasyASP_Date

  Public WeekStarting
  
  Private Sub Class_Initialize()
    '规定周的第一天，可采用下面的值：
    WeekStarting = 2
    '0 = vbUseSystemDayOfWeek - 使用区域语言支持 (NLS) API 设置。 
    '1 = vbSunday - 星期日
    '2 = vbMonday - 星期一
    '3 = vbTuesday - 星期二 
    '4 = vbWednesday - 星期三 
    '5 = vbThursday - 星期四 
    '6 = vbFriday - 星期五 
    '7 = vbSaturday - 星期六
  End Sub
  Private Sub Class_Terminate()
    
  End Sub
  
  '格式化日期时间
  Public Function Format(ByVal iTime, ByVal iFormat)
    If Easp.IsN(iTime) Or Not IsDate(iTime) Then Format = "" : Exit Function
    '调用系统函数格式化时间
    If Instr(",0,1,2,3,4,",","&iFormat&",")>0 Then Format = FormatDateTime(iTime,iFormat) : Exit Function
    Dim diffs,diffd,diffw,diffm,diffy,dire,before,pastTime
    Dim iYear, iMonth, iDay, iHour, iMinute, iSecond,iWeek,tWeek
    Dim iiYear, iiMonth, iiDay, iiHour, iiMinute, iiSecond,iiWeek
    Dim iiiWeek, iiiMonth, iiiiMonth
    Dim SpecialText, SpecialTextRe,i,t
    '取日期时间的值
    iYear = right(Year(iTime),2) : iMonth = Month(iTime) : iDay = Day(iTime)
    iHour = Hour(iTime) : iMinute = Minute(iTime) : iSecond = Second(iTime)
    iiYear = Year(iTime) : iiMonth = right("0"&Month(iTime),2)
    iiDay = right("0"&Day(iTime),2) : iiHour = right("0"&Hour(iTime),2)
    iiMinute = right("0"&Minute(iTime),2) : iiSecond = right("0"&Second(iTime),2)
    tWeek = Weekday(iTime,1)-1 : iWeek = Array("日","一","二","三","四","五","六")
    '如果第二个参数为空或为日期值，则比较时间差
    If isDate(iFormat) Or Easp.IsN(iFormat) Then
      '如果第二个参数为空，则设定为和现在时间相比较
      If Easp.IsN(iFormat) Then : iFormat = Now() : pastTime = true : End If
      dire = Easp.Lang("date-after") : If DateDiff("s",iFormat,iTime)<0 Then : dire = Easp.Lang("date-ago") : before = True : End If
      diffs = Abs(DateDiff("s",iFormat,iTime))
      diffd = Abs(DateDiff("d",iFormat,iTime))
      diffw = Abs(DateDiff("ww",iFormat,iTime))
      diffm = Abs(DateDiff("m",iFormat,iTime))
      diffy = Abs(DateDiff("yyyy",iFormat,iTime))
      If diffs < 60 Then Format = Easp.Lang("date-justnow") : Exit Function
      If diffs < 1800 Then Format = Int(diffs\60) & Easp.Lang("date-minutes")  & dire : Exit Function
      If diffs < 2400 Then Format = Easp.Lang("date-halfhour")  & dire : Exit Function
      If diffs < 3600 Then Format = Int(diffs\60) & Easp.Lang("date-minutes")  & dire : Exit Function
      If diffs < 259200 Then
        If diffd = 3 Then Format = Easp.Lang("date-3days") & dire & " " & iiHour & ":" & iiMinute : Exit Function
        If diffd = 2 Then Format = Easp.IIF(before,Easp.Lang("date-daybeforeyesterday"), Easp.Lang("date-dayaftertomorrow")) & iiHour & ":" & iiMinute : Exit Function
        If diffd = 1 Then Format = Easp.IIF(before,Easp.Lang("date-yesterday"),Easp.Lang("date-tomorrow")) & iiHour & ":" & iiMinute : Exit Function
        Format = Int(diffs\3600) & Easp.Lang("date-hours") & dire : Exit Function
      End If
      If diffd < 7 Then Format = diffd & Easp.Lang("date-days")  & dire & " " & iiHour & ":" & iiMinute : Exit Function
      '如果第二个参数为空，则只显示2周内的相差时间
      If diffd < 14 Then
        If diffw = 1 Then Format = Easp.IIF(before,Easp.Lang("date-lastweek"),Easp.Lang("date-nextweek")) & iWeek(tWeek) & " " & iiHour & ":" & iiMinute : Exit Function
        If Not pastTime Then Format = diffd & Easp.Lang("date-days") & dire : Exit Function
      End If
      '如果第二个参数为具体时间，则显示3年内的相差时间
      If Not pastTime Then
        If diffd < 31 Then
          If diffm = 2 Then Format = Easp.Lang("date-2months") & dire : Exit Function
          If diffm = 1 Then Format = Easp.IIF(before,Easp.Lang("date-lastmonth"),Easp.Lang("date-nextmonth")) & iDay & Easp.Lang("date-day") : Exit Function
          Format = diffw & Easp.Lang("date-weeks") & dire : Exit Function
        End If
        If diffm < 36 Then
          If diffy = 3 Then Format = Easp.Lang("date-3years") & dire : Exit Function
          If diffy = 2 Then Format = Easp.IIF(before,Easp.Lang("date-yearbeforelast"),Easp.Lang("date-yearafternext")) & iMonth & Easp.Lang("date-month") : Exit Function
          If diffy = 1 Then Format = Easp.IIF(before,Easp.Lang("date-last"),Easp.Lang("date-next")) & iMonth & Easp.Lang("date-month") : Exit Function
          Format = diffm & Easp.Lang("date-months") & dire : Exit Function
        End If
        Format = diffy & Easp.Lang("date-years") & dire : Exit Function
      Else
        '如时间超过上述范围则直接显示
        iFormat = "yyyy-mm-dd hh:ii"
      End If
    End If
    iiWeek = Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
    iiiWeek = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
    iiiMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
    iiiiMonth = Array("January","February","March","April","May","June","July","August","September","October","November","December")
    SpecialText = Array("y","m","d","h","i","s","w")
    SpecialTextRe = Array(Chr(0),Chr(1),Chr(2),Chr(3),Chr(4),Chr(5),Chr(6))
    For i = 0 To 6 : iFormat = Replace(iFormat,"\"&SpecialText(i), SpecialTextRe(i)) : Next
    t = Replace(iFormat,"yyyy", iiYear) : t = Replace(t, "yyy", iiYear)
    t = Replace(t, "yy", iYear) : t = Replace(t, "y", iiYear)
    t = Replace(t, "mmmm", Replace(iiiiMonth(iMonth-1),"m",Chr(1))) : t = Replace(t, "mmm", iiiMonth(iMonth-1))
    t = Replace(t, "mm", iiMonth) : t = Replace(t, "m", iMonth)
    t = Replace(t, "dd", iiDay) : t = Replace(t, "d", iDay)
    t = Replace(t, "hh", iiHour) : t = Replace(t, "h", iHour)
    t = Replace(t, "ii", iiMinute) : t = Replace(t, "i", iMinute)
    t = Replace(t, "ss", iiSecond) : t = Replace(t, "s", iSecond)
    t = Replace(t, "www", iiiWeek(tWeek)) : t = Replace(t, "ww", iiWeek(tWeek))
    t = Replace(t, "w", iWeek(tWeek))
    For i = 0 To 6 : t = Replace(t, SpecialTextRe(i),SpecialText(i)) : Next
    Format = t
  End Function

  '取所在月份的第一天
  Public Function FirstDayOfMonth(ByVal d)
    FirstDayOfMonth = CDate(Year(d)&"-"&Month(d)&"-1" & Format(d, " hh:ii:ss"))
  End Function
  
  '取所在月份的最后一天
  Public Function LastDayOfMonth(ByVal d)
    LastDayOfMonth = CDate(DateAdd("d",-1,DateAdd("m",1,Year(d)&"-"&Month(d)&"-1")) & Format(d, " hh:ii:ss"))
  End Function

  '取所在周的第N(1-7)天
  Public Function DayOfWeek(ByVal d, ByVal n)
    DayOfWeek = DateAdd("d",n-Weekday(d,WeekStarting),d)
  End Function
  
  '取所在周的第一天
  Public Function FirstDayOfWeek(ByVal d)
    FirstDayOfWeek = DayOfWeek(d,1)
  End Function
  
  '取所在周的最后一天
  Public Function LastDayOfWeek(ByVal d)
    LastDayOfWeek = DayOfWeek(d,7)
  End Function

  '日期到时间戳函数
  Public Function ToUnixTime(ByRef dateTime, ByRef timeZone)
    If Easp.IsN(dateTime) or Not IsDate(dateTime) Then dateTime = Now
    If Easp.IsN(timeZone) or Not isNumeric(timeZone) Then TimeZone = 0
    ToUnixTime = DateAdd("h", -TimeZone, dateTime)
    ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", ToUnixTime)
  End Function
  '取中国时区时间戳
  Public Function ToUnixTimeCn(ByRef dateTime)
    ToUnixTimeCn = ToUnixTime(dateTime, +8)
  End Function
  '取当前时间戳
  Public Function GetTimeStamp()
    GetTimeStamp = ToUnixTime(Now(), +8)
  End Function
  '时间戳到日期
  Public Function FromUnixTime(ByRef timeStamp, ByRef timeZone)
    If Easp.IsN(timeStamp) Or Not IsNumeric(timeStamp) Then
      FromUnixTime = Now()
      Exit Function
    End If
    If IsEmpty(timeStamp) Or Not IsNumeric(timeZone) Then timeZone = 0
    FromUnixTime = DateAdd("s", timeStamp, "1970-1-1 0:0:0")
    FromUnixTime = DateAdd("h", timeZone, FromUnixTime)
  End Function
  '中国时区时间戳到日期
  Public Function FromUnixTimeCn(ByRef timeStamp)
    FromUnixTimeCn = FromUnixTime(timeStamp, +8)
  End Function
End Class
%>