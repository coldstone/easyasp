<%
'######################################################################
'## easp.log.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Log Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2015-08-03
'## Description :   Log file generator
'##
'######################################################################

Class EasyASP_Log

  Private b_enabled, b_file, b_db, b_inited, _
          s_path, s_siteFolder, s_logFolder, s_rolling, s_file, s_id, _
          i_timer, i_lastTimer, _
          dic, dic_add, dic_style
  
  Private Sub Class_Initialize()
    i_timer = Easp_Timer
    b_enabled = False
  End Sub
  Private Sub Class_Terminate()
    If IsObject(dic) Then Set dic = Nothing
    If IsObject(dic_add) Then Set dic_add = Nothing
    If IsObject(dic_style) Then Set dic_style = Nothing
  End Sub

  Private Sub init()
    If Not b_inited Then
      b_file = True
      b_db = False
      b_inited = False
      s_id = ""
      s_path = "/../"
      Set dic_style = Easp.Json.NewObject
      dic_style("info") = "[{date:Dy-mm-dd hh:ii:ss}, {ip}] ({method} {url}, {run}ms) {msg}"
      dic_style("warn") = "[{date:Dy-mm-dd hh:ii:ss}, {ip}] ({method} {url}, {run}ms):\n  {ua}\n  {msg}"
      dic_style("error") = "[{date:Dy-mm-dd hh:ii:ss}] ({method} {url}, {run}ms)\n  {fn}\n  {msg}"
      s_siteFolder = Easp.Str.GetValueRev(Easp.Fso.MapPath("/"), "\")
      s_logFolder = s_siteFolder & "_log/"
      s_rolling = "day"
      s_file = makeFileName("ymmdd")
      Set dic = Easp.Json.NewObject()
      dic("ip") = Easp.GetIp()
      dic("url") = Easp.GetUrl("")
      dic("method") = Request.ServerVariables("REQUEST_METHOD")
      dic("ua") = Request.ServerVariables("HTTP_USER_AGENT")
      b_inited = True
    End If
  End Sub

  Public Property Let Enable(ByVal bool)
    b_enabled = bool
  End Property

  Public Property Get Enable
    Enable = b_enabled
  End Property

  Private Function makeFileName(Byval fileStyle)
    makeFileName = Year(Now) & "_" & Right("0" & Month(Now), 2) & "/" & s_siteFolder & _
                   Easp.IfThen(Easp.Has(s_id), "_" & s_id) & _
                   "_" & Easp.Date.Format(Now(), fileStyle)
  End Function

  Public Property Let Appender(ByVal string)
    init()
    Dim t, i, ti
    b_file = False
    b_db = False
    b_console = False
    t = Split(string, ",")
    For i = 0 To Ubound(t)
      ti = Trim(t(i))
      If Easp.Str.IsSame(ti, "file") Then
        b_file = True
      ElseIf Easp.Str.IsSame(ti, "db") Then
        b_db = True
      End If
    Next
  End Property

  Public Property Get Appender
    init()
    Dim a : Set a = Easp.Json.NewArray()
    If b_file Then a.Add "file"
    If b_db Then a.Add "db"
    Appender = Join(a.GetArray, ", ")
    Set a = Nothing
  End Property

  Public Property Let FileRolling(ByVal roll)
    init()
    s_rolling = roll
    Select Case roll
      Case "h", "hour"
        s_file = makeFileName("ymmddhh")
      Case "m", "min", "minute"
        s_file = makeFileName("ymmddhhii")
      Case Else
        s_file = makeFileName("ymmdd")
        s_rolling = "day"
    End Select
  End Property

  Public Property Get FileRolling
    init()
    FileRolling = s_rolling
  End Property

  Public Property Let ID(ByVal s)
    init()
    s_id = s
  End Property

  Public Property Get ID
    ID = s_id
  End Property

  Public Property Let SavePath(ByVal string)
    init()
    s_path = string
  End Property

  Public Property Get SavePath
    init()
    SavePath = Easp.Fso.MapPath(s_path) & "\" & s_siteFolder
  End Property

  Public Property Let Style(Byval t, ByVal string)
    init()
    dic_style(t) = string
  End Property
  Public Property Get Style(ByVal t)
    init()
    Style = dic_style(t)
  End Property

  Public Sub [Set](ByVal key, ByVal value)
    init()
    dic.Put key, value
    If Not IsObject(dic_add) Or TypeName(dic_add) = "Nothing" Then
      Set dic_add = dic.Clone()
    Else
      dic_add.Put key, value
    End If
  End Sub

  Public Sub SetOne(ByVal key, ByVal value)
    init()
    If Not IsObject(dic_add) Or TypeName(dic_add) = "Nothing" Then
      Set dic_add = dic.Clone()
    Else
      dic_add.Put key, value
    End If
  End Sub

  Public Sub Start()
    init()
    i_timer = Timer()
  End Sub

  Private Sub makeFile(ByVal t, ByVal msg, ByVal source)
    Dim fileName, header, txt, line
    If Not IsObject(dic_add) Or TypeName(dic_add) = "Nothing" Then Set dic_add = dic.Clone()
    dic_add("date") = Now()
    dic_add("msg") = msg
    dic_add("run") = CInt((Timer() - i_timer) * 1000)
    If Instr(dic_style(t), "\n")>0 Then dic_style(t) = Replace(dic_style(t), "\n", vbCrLf)
    If Easp.Has(source) Then
      line = Easp.Str.GetValueRev(source, ":")
      If IsNumeric(line) Then
        dic_add("fn") = Easp.Str.GetNameRev(source, ":") & ", line " & line
      Else
        dic_add("fn") = source
      End If
    End If
    txt = Easp.Str.Format(dic_style(t), dic_add)
    Set dic_add = Nothing
    fileName = s_path & s_logFolder & s_file & "_" & Lcase(t) & ".log"
    If Easp.Fso.IsFile(fileName) Then
      Call Easp.Fso.AppendFile(fileName, txt & vbCrLf)
    Else
      header = "# " & t & " logs for " & s_siteFolder & vbCrLf
      Call Easp.Fso.CreateFile(fileName, header & txt & vbCrLf)
    End If
  End Sub

  Public Sub Info(ByVal string)
    init()
    If b_file Then Call makeFile("Info", string, Null)
  End Sub

  Public Sub Warn(ByVal string)
    init()
    If b_file Then Call makeFile("Warn", string, Null)
  End Sub

  Public Sub Error(ByVal string, ByVal source)
    init()
    If b_file Then Call makeFile("Error", string, source)
  End Sub

End Class
%>