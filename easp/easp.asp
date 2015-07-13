<%
Option Explicit
'######################################################################
'## easp.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Class
'## Version     :   v2.2 alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014/1/23 09:54
'## Description :   EasyAsp Class
'##
'######################################################################
Dim Easp_Timer : Easp_Timer = Timer()
Dim Easp_DbQueryTimes : Easp_DbQueryTimes = 0
Dim Easp : Set Easp = New EasyAsp : Easp.Init()
Dim EasyAsp_s_html
%>
<!--#include file="easp.config.asp"-->
<%
Class EasyAsp
	Public db,Fso,Upload,Tpl,Aes,[Error],Json,Cache,Http,List,Xml
	Private s_path, s_plugin, s_fsoName, s_dicName, s_charset, s_rq, s_bom
	Private s_url, s_rwtS, s_rwtU, s_cores
	Private o_md5, o_rwt, o_ext, o_regex, o_fso
	Private b_cooen, b_debug, i_rule, i_newId
	Private Sub Class_Initialize()
		s_path		= "/easp/"
		s_plugin	= s_path & "plugin/"
		s_fsoName	= "Scripting.FileSystemObject"
		s_dicName	= "Scripting.Dictionary"
		s_charset	= "UTF-8"
		On Error Resume Next
		Session.CodePage = CharCode(s_charset)
		Response.Charset = s_charset
		Err.Clear
		On Error Goto 0
		s_bom		= "remove"
		s_rq		= Request.QueryString()
		i_rule		= 1
		b_cooen		= False
		b_debug		= False
		i_newId		= 0
		Set o_rwt 	= Server.CreateObject(s_dicName)
		Set o_ext 	= Server.CreateObject(s_dicName)
		Set o_regex = New Regexp
		o_regex.Global = True
		o_regex.IgnoreCase = True
		s_cores		= "[Error],db,List,o_md5,Fso,Upload,Tpl,Aes,JSON,Cache,Http,Xml"
		Core_Do "on", s_cores
	End Sub
	Private Sub Class_Terminate()
		Core_Do "off", s_cores
		ClearExt() : Set o_ext	= Nothing
		Set o_rwt	= Nothing
		Set o_regex = Nothing
		If isObject(o_fso) Then Set o_fso = Nothing
	End Sub
	Public Sub Init()
		Set [error] = New EasyAsp_Error
		[error](1) = "包含文件内部运行错误，请检查包含文件代码！"
		[error](2) = "读取文件错误，文件未找到！"
		[error](3) = "EasyASP系统路径错误，请检查'easp.config.asp'文件中的 Easp.BasePath 设置！"
		Set db = New EasyAsp_db
'{init}
		Set List = New EasyAsp_List
		Set Fso = New EasyAsp_Fso
		Set Upload = New EasyAsp_Upload
		Set Tpl = New EasyAsp_Tpl
		Set Aes = New EasyAsp_AES
		Set Json = New EasyAsp_JSON
		Set Cache = New EasyAsp_Cache
		Set Http = New EasyAsp_Http
		Set Xml = New EasyAsp_Xml
'{/init}
	End Sub
	Private Sub Core_Do(ByVal t, ByVal s)
		Dim a_core, i : a_core = Split(s,",")
		Select Case t
			Case "on"
				For i = 0 To Ubound(a_core)
					Execute "Set " & a_core(i) & " = New EasyAsp_obj"
				Next
			Case "off"
				For i = Ubound(a_core) To 0 Step -1
					Execute "Set " & a_core(i) & " = Nothing"
				Next
		End Select
	End Sub
	'设置和读取Easp路径配置
	Public Property Let basePath(ByVal p)
		s_path = FixAbsPath(p)
		s_plugin = s_path & "plugin/"
	End Property
	Public Property Get basePath()
		basePath = s_path
	End Property
	'设置和读取Easp插件文件夹路径配置
	Public Property Let pluginPath(ByVal p)
		s_plugin = FixAbsPath(p)
	End Property
	Public Property Get pluginPath()
		pluginPath = s_plugin
	End Property
	'设置和读取FSO组件名称
	Public Property Let fsoName(ByVal s)
		s_fsoName = s
	End Property
	Public Property Get fsoName()
		fsoName = s_fsoName
	End Property
	'设置和读取
	Public Property Let [CharSet](ByVal s)
		s_charset = Ucase(s)
		On Error Resume Next
		Session.CodePage = CharCode(s_charset)
		Response.Charset = s_charset
		Err.Clear
		On Error Goto 0
	End Property
	Public Property Get [CharSet]()
		[CharSet] = s_charset
	End Property
	'设置和读取如何处理UTF-8文件的BOM信息
	Public Property Let FileBOM(ByVal s)
		s_bom = Lcase(s)
	End Property
	Public Property Get FileBOM()
		FileBOM = s_bom
	End Property
	'设置和读取是否加密Cookies信息
	Public Property Let CookieEncode(ByVal b)
		b_cooen = b
	End Property
	Public Property Get CookieEncode()
		CookieEncode = b_cooen
	End Property
	'设置和读取是否开启调试模式
	Public Property Let [Debug](ByVal b)
		b_debug = b
		[error].debug = b
	End Property
	Public Property Get [Debug]
		[Debug] = b_debug
	End Property
	'读取本页查询数据库次数
	Public Property Get DbQueryTimes
		DbQueryTimes = Easp_DbQueryTimes
	End Property
	'读取服务器端执行时间（秒）
	Public Property Get ScriptTime
		ScriptTime = toNumber(GetScriptTime(0)/1000,3)
	End Property

	Private Function FixAbsPath(ByVal p)
		p = IIF(Left(p,1)= "/", p, "/" & p)
		p = IIF(Right(p,1)="/", p, p & "/")
		FixAbsPath = p
	End Function
	Private Function rqsv(ByVal s)
		rqsv = Request.ServerVariables(s)
	End Function
	
	'输出字符串(简易断点调试)
	Sub W(ByVal s) : Response.Write(s) : End Sub
	Sub WC(ByVal s) : W(s & VbCrLf) : End Sub
	Sub WF(ByVal s) : W s : Response.Flush() : End Sub
	Sub WN(ByVal s) : W s & "<br />" & VbCrLf : End Sub
	Sub WNF(ByVal s) : WN s : Response.Flush() : End Sub
	Sub WNH(ByVal s) : WN HtmlEncode(s) : End Sub
	Sub WE(ByVal s) : W s : Set Easp = Nothing : Response.End() : End Sub
	Sub WT(ByVal s) : W "<span style=""background-color:#FF3;color:blue;margin-right:10px;"">"&s&"</span>" : End Sub
	'格式化字符串（首下标为1）
	Function Str(ByVal s, ByVal v)
		Str = FormatString(s, v, 1)
	End Function
	'格式化字符串（首下标为0）
	Function Format(ByVal s, ByVal v)
		Format = FormatString(s, v, 0)
	End Function
	Private Function FormatString(ByVal s, ByRef v, ByVal t)
		Dim i,n,k
		s = Replace(s,"\\",Chr(0))
		s = Replace(s,"\{",Chr(1))
		Select Case VarType(v)
			'数组
			Case 8192,8194,8204,8209
				For i = 0 To Ubound(v)
					s = FormatReplace(s,i+t,v(i))
				Next
			'对象
			Case 9
				Select Case TypeName(v)
					'记录集
					Case "Recordset"
						For i = 0 To v.Fields.Count - 1
							s = FormatReplace(s,i+t,v(i))
							s = FormatReplace(s,v.Fields.Item(i+t).Name,v(i))
						Next
					'字典
					Case "Dictionary"
						For Each k In v
							s = FormatReplace(s,k,v(k))
						Next
					'Easp List
					Case "EasyAsp_List"
						For i = 0 To v.End
							s = FormatReplace(s,i+t,v(i))
							s = FormatReplace(s,v.IndexHash(i),v(i))
						Next
					'正则搜索子集合
					Case "ISubMatches", "SubMatches"
						For i = 0 To v.Count - 1
							s = FormatReplace(s,i+t,v(i))
						Next
				End Select
			'字符串
			Case 8
				Select Case TypeName(v)
					'正则搜索集合
					Case "IMatch2", "Match"
						s = FormatReplace(s,t,v.Value)
						For i = 0 To v.SubMatches.Count - 1
							s = FormatReplace(s,i+t+1,v.SubMatches(i))
						Next
					'字符串
					Case Else
						s = FormatReplace(s,t,v)
				End Select
			Case Else
				s = FormatReplace(s,t,v)
		End Select
		s = Replace(s,Chr(1),"{")
		FormatString = Replace(s,Chr(0),"\")
	End Function
	'格式化Format内标签参数
	Private Function FormatReplace(ByVal s, ByVal t, ByVal v)
		Dim tmp,rule,ru,kind,matches,match
		v = IfHas(v,"")
		rule = "\{" & t & "(:((N[,\(%]?(\d+)?)|(D[^\}]+)|(E[^\}]+)|U|L|\d+([^\}]+)?))\}"
		If Test(s,rule) Then
			Set matches = RegMatch(s,rule)
			For Each match In matches
				kind = RegReplace(match.Value, rule, "$2")
				ru = match.Value
				Select Case Left(kind,1)
					'截取字符串
					Case "1","2","3","4","5","6","7","8","9"
						s = Replace(s, ru, CutStr(v,regReplace(kind,"^(\d+)(.+)?$","$1:$2")))
					'数字
					Case "N"
						If isNumeric(v) Then
							Dim style,group,parens,percent,deci
							style = RegReplace(kind,"^N([,\(%])?(\d+)?$","$1")
							If style = "," Then group = -1
							If style = "(" Then parens = -1
							If style = "%" Then percent = -1
							deci = RegReplace(kind,"^N([,\(%])?(\d+)?$","$2")
							If IsN(style) And IsN(deci) Then
								s = Replace(s, ru, IIF(Instr(Cstr(v),".")>0 And v<1,"0" & v,v),1,-1,1)
							Else
								deci = IfHas(deci,-1)
								If percent Then
									s = Replace(s, ru, FormatNumber(v*100,deci,-1) & "%",1,-1,1)
								Else
									s = Replace(s, ru, FormatNumber(v,deci,-1,parens,group),1,-1,1)
								End If
							End If
						End If
					'日期
					Case "D"
						If isDate(v) Then
							s = Replace(s, ru, DateTime(v,Mid(kind,2)),1,-1,1)
						End If
					'转大写
					Case "U"
						s = Replace(s, ru, UCase(v),1,-1,1)
					'转小写
					Case "L"
						s = Replace(s, ru, LCase(v),1,-1,1)
					'表达式
					Case "E"
						tmp = Replace(Mid(kind,2),"%s", "v")
						tmp = Eval(tmp)
						s = Replace(s, ru, tmp,1,-1,1)
				End Select
			Next
		Else
			s = Replace(s,"{" & t & "}",v,1,-1,1)
		End If
		FormatReplace = s
	End Function
	'服务器端生成唯一不重复编号(10位字符串)
	Public Function NewID()
		Dim id
		If i_newId = 0 Then
			i_newId = (DateDiff("s","1970-01-01",DateTime(Now(),"y-mm-dd"))+Timer())*10000
		End If
		i_newId = i_newId + 1
		id = i_newId
		NewID = NumberToString(id)
	End Function
	Private Function NumberToString(n)
		Dim t(10), v, c, m
		c = 10
		v = n / 36
		Do While v > 0 And c > 0
			c = c - 1
			m = Int(n - Int(n / 36) * 36)
			t(c) = Chr(IIF(m<10, m+48, m+55))
			n = v
			v = n / 36
		Loop
		NumberToString = Join(t,"")
	End Function
	'服务器端跳转
	Sub RR(ByVal u)
		Response.Redirect(u)
	End Sub
	'判断是否为空值
	Function isN(ByVal s)
		isN = False
		Select Case VarType(s)
			Case vbEmpty, vbNull
				isN = True : Exit Function
			Case vbString
				If s="" Then isN = True : Exit Function
			Case vbObject
				Select Case TypeName(s)
					Case "Nothing","Empty"
						isN = True : Exit Function
					Case "Recordset"
						If s.State = 0 Then isN = True : Exit Function
						If s.Bof And s.Eof Then isN = True : Exit Function
					Case "Dictionary"
						If s.Count = 0 Then isN = True : Exit Function
				End Select
			Case vbArray,8194,8204,8209
				If Ubound(s)=-1 Then isN = True : Exit Function
		End Select
	End Function
	'判断是否不为空值
	Function Has(ByVal s)
		Has = Not isN(s)
	End Function
	'判断三元表达式
	Function IIF(ByVal Cn, ByVal T, ByVal F)
		If Cn Then
			IIF = T
		Else
			IIF = F
		End If
	End Function
	'如果条件成立则返回某值
	Function IfThen(ByVal Cn, ByVal T)
		IfThen = IIF(Cn,T,"")
	End Function
	'如果第1项不为空则返回第1项，否则返回第2项
	Function IfHas(ByVal v1, ByVal v2)
		IfHas = IIF(Has(v1), v1, v2)
	End Function
	'不够长度的字符串填充内容
	Function Fill(ByVal s, ByVal l, ByVal f)
		Dim le,i,tmp
		le = len(s)
		If isN(f) Then f = "&nbsp;"
		If le < l Then
			For i = 1 To l-le
				tmp = tmp & f
			Next
		End If
		Fill = s & tmp
	End Function
	'服务器端输出javascript
	Sub Js(ByVal s)
		W JsCode(s)
	End Sub
	'生成javascript代码
	Function JsCode(ByVal s)
		JsCode = Str("<{1} type=""text/java{1}"">{2}{3}{4}{2}</{1}>{2}", Array("sc"&"ript",vbCrLf,vbTab,s))
	End Function
	'服务器端输出javascript弹出消息框并返回前页
	Sub Alert(ByVal s)
		WE JsCode(Str("alert('{1}');history.go(-1);",JsEncode(s)))
	End Sub
	'服务器端输出javascript弹出消息框并转到URL
	Sub AlertUrl(ByVal s, ByVal u)
		WE JsCode(Str("alert('{1}');location.href='{2}';",Array(JsEncode(s),u)))
	End Sub
	'服务器端输出javascript确认消息框并根据选择转到URL
	Sub ConfirmUrl(ByVal s, ByVal t, ByVal f)
		WE JsCode(Str("location.href=confirm('{1}')?'{2}':'{3}';",Array(JsEncode(s),t,f)))
	End Sub
	'处理字符串中的Javascript特殊字符
	Function JsEncode(ByVal s)
		If isN(s) Then JsEncode = "" : Exit Function
		Dim arr1, arr2, i, j, c, p, t
		arr1 = Array(&h27,&h22,&h5C,&h2F,&h08,&h0C,&h0A,&h0D,&h09)
		arr2 = Array(&h27,&h22,&h5C,&h2F,&h62,&h66,&h6E,&h72,&h749)
		For i = 1 To Len(s)
			p = True
			c = Mid(s, i, 1)
			For j = 0 To Ubound(arr1)
				If c = Chr(arr1(j)) Then
					t = t & "\" & Chr(arr2(j))
					p = False
					Exit For
				End If
			Next
			If p Then 
				Dim a
				a = AscW(c)
				If a > 31 And a < 127 Then
					t = t & c
				ElseIf a > -1 Or a < 65535 Then
					t = t & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
				End If 
			End If
		Next
		JsEncode = t
	End Function
	'转换字符为HTML实体
	Function ASCII(s)
		If IsN(s) Then ASCII = "" : Exit Function
		Dim tmp,i,t
		For i = 1 To len(s)
			t = AscW(Mid(s,i,1))
			If t<0 Then t = t + 65536
			tmp = tmp & "&#" & t & ";"
		Next
		ASCII = tmp
	End Function
	'特殊字符编码
	Function Escape(ByVal ss)
		If isN(ss) Then Escape = "" : Exit Function
		Dim i,c,a,s : s = ""
		For i = 1 To Len(ss)
			c = Mid(ss,i,1)
			a = ASCW(c)
			If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
				s = s & c
			ElseIf InStr("@*_+-./",c)>0 Then
				s = s & c
			ElseIf a>0 and a<16 Then
				s = s & "%0" & Hex(a)
			ElseIf a>=16 and a<256 Then
				s = s & "%" & Hex(a)
			Else
				s = s & "%u" & Hex(a)
			End If
		Next
		Escape = s
	End Function
	'特殊字符解码
	Function UnEscape(ByVal ss)
		If isN(ss) Then UnEscape = "" : Exit Function
		Dim x, s
		x = InStr(ss,"%")
		s = ""
		Do While x>0
			s = s & Mid(ss,1,x-1)
			If LCase(Mid(ss,x+1,1))="u" Then
				s = s & ChrW(CLng("&H"&Mid(ss,x+2,4)))
				ss = Mid(ss,x+6)
			Else
				s = s & Chr(CLng("&H"&Mid(ss,x+1,2)))
				ss = Mid(ss,x+3)
			End If
			x=InStr(ss,"%")
		Loop
		UnEscape = s & ss
	End Function
	'格式化日期时间
	Function DateTime(ByVal iTime, ByVal iFormat)
		If IsN(iTime) Then DateTime = "" : Exit Function
		If Not IsDate(iTime) Then DateTime = "Date Error" : Exit Function
		If Instr(",0,1,2,3,4,",","&iFormat&",")>0 Then DateTime = FormatDateTime(iTime,iFormat) : Exit Function
		Dim diffs,diffd,diffw,diffm,diffy,dire,before,pastTime
		Dim iYear, iMonth, iDay, iHour, iMinute, iSecond,iWeek,tWeek
		Dim iiYear, iiMonth, iiDay, iiHour, iiMinute, iiSecond,iiWeek
		Dim iiiWeek, iiiMonth, iiiiMonth
		Dim SpecialText, SpecialTextRe,i,t
		iYear = right(Year(iTime),2) : iMonth = Month(iTime) : iDay = Day(iTime)
		iHour = Hour(iTime) : iMinute = Minute(iTime) : iSecond = Second(iTime)
		iiYear = Year(iTime) : iiMonth = right("0"&Month(iTime),2)
		iiDay = right("0"&Day(iTime),2) : iiHour = right("0"&Hour(iTime),2)
		iiMinute = right("0"&Minute(iTime),2) : iiSecond = right("0"&Second(iTime),2)
		tWeek = Weekday(iTime)-1 : iWeek = Array("日","一","二","三","四","五","六")
		If isDate(iFormat) or isN(iFormat) Then
			If isN(iFormat) Then : iFormat = Now() : pastTime = true : End If
			dire = "后" : If DateDiff("s",iFormat,iTime)<0 Then : dire = "前" : before = True : End If
			diffs = Abs(DateDiff("s",iFormat,iTime))
			diffd = Abs(DateDiff("d",iFormat,iTime))
			diffw = Abs(DateDiff("ww",iFormat,iTime))
			diffm = Abs(DateDiff("m",iFormat,iTime))
			diffy = Abs(DateDiff("yyyy",iFormat,iTime))
			If diffs < 60 Then DateTime = "刚刚" : Exit Function
			If diffs < 1800 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 2400 Then DateTime = "半小时"  & dire : Exit Function
			If diffs < 3600 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 259200 Then
				If diffd = 3 Then DateTime = "3天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
				If diffd = 2 Then DateTime = IIF(before,"前天 ","后天 ") & iiHour & ":" & iiMinute : Exit Function
				If diffd = 1 Then DateTime = IIF(before,"昨天 ","明天 ") & iiHour & ":" & iiMinute : Exit Function
				DateTime = Int(diffs\3600) & "小时" & dire : Exit Function
			End If
			If diffd < 7 Then DateTime = diffd & "天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
			If diffd < 14 Then
				If diffw = 1 Then DateTime = IIF(before,"上星期","下星期") & iWeek(tWeek) & " " & iiHour & ":" & iiMinute : Exit Function
				If Not pastTime Then DateTime = diffd & "天" & dire : Exit Function
			End If
			If Not pastTime Then
				If diffd < 31 Then
					If diffm = 2 Then DateTime = "2个月" & dire : Exit Function
					If diffm = 1 Then DateTime = IIF(before,"上个月","下个月") & iDay & "日" : Exit Function
					DateTime = diffw & "星期" & dire : Exit Function
				End If
				If diffm < 36 Then
					If diffy = 3 Then DateTime = "3年" & dire : Exit Function
					If diffy = 2 Then DateTime = IIF(before,"前年","后年") & iMonth & "月" : Exit Function
					If diffy = 1 Then DateTime = IIF(before,"去年","明年") & iMonth & "月" : Exit Function
					DateTime = diffm & "个月" & dire : Exit Function
				End If
				DateTime = diffy & "年" & dire : Exit Function
			Else
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
		DateTime = t
	End Function
	'伪Rewrite规则设置（传统方法）
	Sub RewriteRule(ByVal s, ByVal u)
		If (Left(s,3)<>"^\/" And Left(s,2)<>"\/" And Left(s,2)<>"^/" And Left(s,1)<>"/") Or Left(u,1)<>"/" Then Exit Sub
		o_rwt.Add ("rule" & i_rule), Array(s,u)
		i_rule = i_rule + 1
	End Sub
	'伪Rewrite规则设置（推荐方法）
	Sub Rewrite(ByVal p, ByVal s, Byval u)
		Dim rp,arr,i,rs,ru
		If Left(s,1) = "^" Then s = Mid(s,2)
		If Right(s,1) = "$" Then s = Left(s,Len(s)-1)
		If IsN(p) Then p = GetUrl(0)
		arr = Split(p,"|")
		For i = 0 To Ubound(arr)
			rp = RegEncode(arr(i))
			rs = "^" & rp & "\?" & s & "$"
			ru = arr(i) & "?" & u
			RewriteRule rs, ru
			rp = ""
		Next
	End Sub
	'检测本页是否是Rewrite
	Function isRewrite()
		Dim rule,i
		isRewrite = False
		If i_rule = 1 Then Exit Function
		If Has(o_rwt) Then
			s_url = GetUrl(1)
			For Each i In o_rwt
				rule = o_rwt(i)(0)
				If Easp_Test(s_url,rule) Then
					isRewrite = True
					s_rwtS = rule
					s_rwtU = CRight(o_rwt(i)(1),"?")
					Exit For
				End If
			Next
		End If
	End Function
	'替换Url参数
	Function ReplaceUrl(ByVal p, ByVal s)
		Dim arrQs,tmp,i
		If isRewrite Then
			arrQs = Split(s_rwtU,"&")
			For i = 0 To Ubound(arrQs)
				If p = CLeft(arrQs(i),"=") Then
					tmp = CRight(arrQs(i),"=")
					Exit For
				End If
			Next
			If Left(tmp,1) = "$" Then
				ReplaceUrl = ReplacePart(s_url, s_rwtS, tmp, s)
			End If
		Else
			ReplaceUrl = GetUrlWith("-" & p, p & "=" & s)
		End If
	End Function
	'获取QueryString值，支持取Rewrite值
	Function [Get](Byval s)
		Dim tmp, i, arrQs, t
		If Instr(s,":")>0 Then
		'如果有类型参数，则取出为t
			t = CRight(s,":") : s = CLeft(s,":")
		End If
		If isRewrite Then
		'如果是Rewrite的页面地址
			arrQs = Split(s_rwtU,"&")
			For i = 0 To Ubound(arrQs)
				If s = CLeft(arrQs(i),"=") Then
					tmp = RegReplace(s_url,s_rwtS,CRight(arrQs(i),"="))
					Exit For
				End If
			Next
		Else
		'否则直接取QueryString
			tmp = Request.QueryString(s)
		End If
		[Get] = Safe(tmp,t)
	End Function
	'UTF-8编码下取中文参数（仅utf-8下有效，感谢Easp.智者提供此方法）
	Function GetCn(ByVal s)
	Dim tmp
	Session.CodePage = 936
	tmp = [Get](s)
	Session.CodePage = 65001
		GetCn = tmp
	End Function
	'取Form值，包括上传文件时的普通Form值
	Function Post(ByVal s)
		Dim t,tmp : tmp = ""
		If Instr(s,":")>0 Then
			t = CRight(s,":") : s = CLeft(s,":")
		End If
		Dim FormType
		If Has(Request.ServerVariables("HTTP_CONTENT_TYPE")) Then
			FormType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")(0)
		Else
			FormType = "notupload"
		End If
		If LCase(FormType) = "multipart/form-data" Then
			If Lcase(TypeName(upload)) = "easyasp_upload" Then
				If upload.Form.Count>0 Then tmp = upload.Form(s)
			End If
		Else
			tmp = Request.Form(s)
		End If
		Post = Safe(tmp,t)
	End Function
	'获取Request值
	Function R(ByVal s)
		If Has([Get](s)) Then
			R = [Get](s)
		ElseIf Has(Post(s)) Then
			R = Post(s)
		ElseIf Has(Cookie(s)) Then
			R = Cookie(s)
		ElseIf Has(Request.ServerVariables(s)) Then
			R = Request.ServerVariables(s)
		ElseIf Has(Request.ClientCertificate(s)) Then
			R = Request.ClientCertificate(s)
		Else'If IsN(Request(s)) Then
			R = Empty
		'Else
		'	R = Request.BinaryRead(s)
		End If
	End Function
	'安全获取值新版
	Function Safe(ByVal s, ByVal t)
		Dim spl,d,l,li,i,tmp,arr() : l = False
		'如果类型中有默认值
		If Instr(t,":")>0 Then
			d = CRight(t,":") : t = CLeft(t,":")
		End If
		If Instr(",sa,da,na,se,de,ne,", "," & Left(LCase(t),2) & ",")>0 Then
			'如果有分隔符且要警告
			If Len(t)>2 Then
				spl = Mid(t,3) : t = LCase(Left(t,2)) : l = True
			End If
		ElseIf Instr("sdn", Left(LCase(t),1))>0 Then
			'如果有分隔符且不警告
			If Len(t)>1 Then
				spl = Mid(t,2) : t = LCase(Left(t,1)) : l = True
			End If
		ElseIf Has(t) Then
			'仅有分隔符无类型
			spl = t : t = "" : l = True
		End If
		li = Split(s,spl)
		If l Then Redim arr(Ubound(li))
		For i = 0 To Ubound(li)
			If i<>0 Then tmp = tmp & spl
			Select Case t
				Case "s","sa","se"
				'字符串类型
					If isN(li(i)) Then li(i) = d
					tmp = tmp & Replace(li(i),"'","''")
					If l Then arr(i) = Replace(li(i),"'","''")
				Case "d","da","de"
				'日期类型
					If t = "da" Then
						If Not isDate(li(i)) And Has(li(i)) Then Alert("不正确的日期值！")
					ElseIf t = "de" Then
						If Not isDate(li(i)) And Has(li(i)) Then [error].Throw("不正确的日期值！")
					End If
					tmp = IIF(isDate(li(i)), tmp & li(i), tmp & d)
					If l Then arr(i) = IIF(isDate(li(i)), li(i), d)
				Case "n","na","ne"
				'数字类型
					If t = "na" Then
						If Not isNumeric(li(i)) And Has(li(i)) Then Alert("不正确的数值！")
					ElseIf t = "ne" Then
						If Not isNumeric(li(i)) And Has(li(i)) Then [error].Throw("不正确的数值！")
					End If
					tmp = IIF(isNumeric(li(i)), tmp & li(i), tmp & d)
					If l Then arr(i) = IIF(isNumeric(li(i)), li(i), d)
				Case Else
				'未指定类型则不处理
					tmp = IIF(isN(li(i)), tmp & d, tmp & li(i))
					If l Then arr(i) = IIF(isN(li(i)), d, li(i))
			End Select
			If l Then
				If IsN(arr(i)) Then arr(i) = Empty
			End If
		Next
		If IsN(tmp) Then tmp = Empty
		Safe = IIF(l,arr,tmp)
	End Function
	'检查提交数据来源
	Function CheckDataFrom()
		Dim v1, v2
		CheckDataFrom = False
		v1 = Lcase(Cstr(rqsv("HTTP_REFERER")))
		v2 = Lcase(Cstr(rqsv("SERVER_NAME")))
		v1 = Mid(v1,Instr(v1,"://")+3,len(v2))
		If v1 = v2 Then
			CheckDataFrom = True
		End If
	end Function
	'检查提交数据来源，不是本站则弹出警告框
	Sub CheckDataFromA()
		If Not CheckDataFrom Then Alert "禁止从站点外部提交数据！"
	end Sub
	'截取长字符串左边部分并以特殊符号代替
	Function CutStr(ByVal s, ByVal strlen)
		If IsN(s) Then CutStr = "" : Exit Function
		If IsN(strlen) or strlen = "0" Then CutStr = s : Exit Function
		Dim l,t,i,j,d,f,n
		s = Replace(s,vbCrLf,"")
		s = Replace(s,vbTab,"")
		l = len(s) : t = 0 : d = ChrW(8230) : f = Easp_Param(strlen)
		If Instr(strlen,":")>0 Then
			d = IfHas(f(1),"")
		End If
		strlen = Int(f(0)) : f = "" : n = 0
		For j = 1 To Len(d)
			n = IIF(Abs(Ascw(Mid(d,j,1)))>255, n+2, n+1)
		Next
		strlen = strlen - n
		For i = 1 to l
			t = IIF(Abs(Ascw(Mid(s,i,1)))>255, t+2, t+1)
			If t >= strlen Then
				f = Left(s,i) & d
				Exit For
			Else
				f = s
			End If
		Next
		CutStr = f
	End Function
	'取字符隔开的左段
	Function CLeft(ByVal s, ByVal m)
		CLeft = Easp_LR(s,m,0)
	End Function
	'取字符隔开的右段
	Function CRight(ByVal s, ByVal m)
		CRight = Easp_LR(s,m,1)
	End Function
	'获取当前文件的地址
	Function GetUrl(param)
		Dim script_name,url,dir
		Dim out,qitem,qtemp,i,hasQS,qstring
		script_name = rqsv("SCRIPT_NAME")
		url = script_name
		dir  = Left(script_name,InstrRev(script_name,"/"))
		If isN(param) or param = "-1" Then
			Dim ustart,uport
			If rqsv("HTTPS")="on" Then
				ustart = "https://"
				uport = IIF(Int(rqsv("SERVER_PORT"))=443,"",":"&rqsv("SERVER_PORT"))
			Else
				ustart = "http://"
				uport = IIF(Int(rqsv("SERVER_PORT"))=80,"",":"&rqsv("SERVER_PORT"))
			End If
			url = ustart & rqsv("SERVER_NAME") & uport
			If isN(param) Then
				url = url & script_name
			Else
				GetUrl = url : Exit Function
			End If
			If Has(s_rq) Then url = url & "?" & s_rq
			GetUrl = url : Exit Function
		End If
		If param = "0" Then : GetUrl = url : Exit Function
		If param = "2" Then : GetUrl = dir : Exit Function
		If InStr(param,":")>0 Then
			url = dir
			out = Mid(param,2)
			hasQS = IIF(isN(out),0,1)
		Else
			out = param : hasQS = 1
		End If
		If Has(s_rq) Then
			If param="1" Or hasQS = 0 Then
				url = url & "?" & s_rq
			Else
				qtemp = "" : i = 0 : out = ","&out&","
				qstring = IIF(InStr(out,"-")>0,"Not InStr(out,"",-""&qitem&"","")>0","InStr(out,"",""&qitem&"","")>0")
				For Each qitem In Request.QueryString()
					If Eval(qstring) Then
						If i<>0 Then qtemp = qtemp & "&"
						qtemp = qtemp & qitem & "=" & Request.QueryString(qitem)
						i = i + 1
					End If
				Next
				If Has(qtemp) Then url = url & "?" & qtemp
			End If
		End If
		GetUrl = url
	End Function
	'获取本页URL参数并带上新的URL参数
	Function GetUrlWith(ByVal p, ByVal v)
		Dim u,s,n
		s = IIF(p=-1,GetUrl(-1)&"/","")
		s = IIF(IsN(p),GetUrl(""),GetUrl(0))
		If Instr(p,":")>0 Then
			If Has(CLeft(p,":")) Then
				n = Cleft(p,":") : p = CRight(p,":")
			End If
		End If
		u = GetUrl(p)
		If Left(p,1)=":" Then s = Left(u,InstrRev(u,"/"))
		u = u & IfThen(Has(v),IIF(isN(Mid(u,len(s)+1)),"?","&") & v)
		If Has(n) Then
			If Instr(u,"?")>0 Then
				u = n & Mid(u,Instr(u,"?"))
			Else
				u = n
			End If
		End If
		GetUrlWith = u
	End Function
	'获取用户IP地址
	Function GetIP()
		Dim addr, x, y
		x = rqsv("HTTP_X_FORWARDED_FOR")
		y = rqsv("REMOTE_ADDR")
		addr = IIF(isN(x) or lCase(x)="unknown",y,x)
		If InStr(addr,".")=0 Then addr = "0.0.0.0"
		GetIP = addr
	End Function
	'仅格式化HTML文本中的空格和换行
	Function HtmlFormat(ByVal s)
		If Has(s) Then
			Dim m,Match : Set m = RegMatch(s, "<([^>]+)>")
			For Each Match In m
				 s = Replace(s, Match.SubMatches(0), regReplace(Match.SubMatches(0), "\s+", Chr(0)))
			Next
			Set m = Nothing
			s = Replace(s, Chr(32), "&nbsp;")
			s = Replace(s, Chr(9), "&nbsp;&nbsp; &nbsp;")
			s = Replace(s, Chr(0), " ")
			s = regReplace(s, "(<[^>]+>)\s+", "$1")
			s = Replace(s, vbCrLf, "<br />")
		End If
		HtmlFormat = s
	End Function
	'将HTML代码转换为文本实体
	Function HtmlEncode(ByVal s)
		If Has(s) Then
			s = Replace(s, Chr(38), "&#38;")
			s = Replace(s, "<", "&lt;")
			s = Replace(s, ">", "&gt;")
			s = Replace(s, Chr(39), "&#39;")
			s = Replace(s, Chr(32), "&nbsp;")
			s = Replace(s, Chr(34), "&quot;")
			s = Replace(s, Chr(9), "&nbsp;&nbsp; &nbsp;")
			s = Replace(s, vbCrLf, "<br />")
		End If
		HtmlEncode = s
	End Function
	'将HTML文本转换为HTML代码
	Function HtmlDecode(ByVal s)
		If Has(s) Then
			s = regReplace(s, "<br\s*/?\s*>", vbCrLf)
			s = Replace(s, "&nbsp;&nbsp; &nbsp;", Chr(9))
			s = Replace(s, "&quot;", Chr(34))
			s = Replace(s, "&nbsp;", Chr(32))
			s = Replace(s, "&#39;", Chr(39))
			s = Replace(s, "&apos;", Chr(39))
			s = Replace(s, "&gt;", ">")
			s = Replace(s, "&lt;", "<")
			s = Replace(s, "&amp;", Chr(38))
			s = Replace(s, "&#38;", Chr(38))
		End If
		HtmlDecode = s
	End Function
	'过滤HTML标签
	Function HtmlFilter(ByVal s)
		If Has(s) Then
			If IsN(s) Then HtmlFilter = "" : Exit Function
			s = regReplace(s,"<[^>]+>","")
			s = Replace(s, ">", "&gt;")
			s = Replace(s, "<", "&lt;")
		End If
		HtmlFilter = s
	End Function
	'精确到毫秒的脚本执行时间
	Function GetScriptTime(t)
		If t = "" Or t = "0" Then t = Easp_Timer
		GetScriptTime = FormatNumber((Timer()-t)*1000, 2, -1)
	End Function
	'取指定长度的随机字符串
	Function RandStr(ByVal cfg)
		Dim a, p, l, t, reg, m, mi, ma
		cfg = Replace(Replace(Replace(cfg,"\<",Chr(0)),"\>",Chr(1)),"\:",Chr(2))
		a = ""
		If Easp_Test(cfg, "(<\d+>|<\d+-\d+>)") Then
			t = cfg
			p = Easp_Param(cfg)
			If Not isN(p(1)) Then
				a = p(1) : t = p(0) : p = ""
			End If
			Set reg = RegMatch(cfg, "(<\d+>|<\d+-\d+>)")
			For Each m In reg
				p = m.SubMatches(0)
				l = Mid(p,2,Len(p)-2)
				If Easp_Test(l,"^\d+$") Then
					t = Replace(t,p,Easp_RandString(l,a),1,1)
				Else
					mi = CLeft(l,"-")
					ma = CRight(l,"-")
					t =  Replace(t,p,Rand(mi, ma),1,1)
				End If
			Next
			Set reg = Nothing
		ElseIf Easp_Test(cfg,"^\d+-\d+$") Then
			mi = CLeft(cfg,"-")
			ma = CRight(cfg,"-")
			t = Rand(mi, ma)
		ElseIf Easp_Test(cfg, "^(\d+)|(\d+:.)$") Then
			l = cfg : p = Easp_Param(cfg)
			If Not isN(p(1)) Then
				a = p(1) : l = p(0) : p = ""
			End If
			t = Easp_RandString(l, a)
		Else
			t = cfg
		End If
		RandStr = Replace(Replace(Replace(t,Chr(0),"<"),Chr(1),">"),Chr(2),":")
	End Function
	Private Function Easp_RandString(ByVal length, ByVal allowStr)
		Dim i
		If IsN(allowStr) Then allowStr = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
		For i = 1 To length
			Randomize(Timer) : Easp_RandString = Easp_RandString & Mid(allowStr, Int(Len(allowStr) * Rnd + 1), 1)
		Next
	End Function
	'取一个随机数
	Function Rand(ByVal min, ByVal max)
		Randomize(Timer) : Rand = Int((max - min + 1) * Rnd + min)
	End Function
	'格式化数字
	Function toNumber(ByVal n, ByVal d)
		toNumber = FormatNumber(n,d,-1)
	End Function
	'将数字转换为货币格式
	Function toPrice(ByVal n)
		toPrice = FormatCurrency(n,2,-1,0,-1)
	End Function
	'将数字转换为百分比格式
	Function toPercent(ByVal n)
		toPercent = FormatPercent(n,2,-1)
	End Function
	'关闭对象并释放资源
	Sub C(ByRef o)
		On Error Resume Next
		o.Close() : Set o = Nothing
		Err.Clear()
	End Sub
	'不缓存页面信息
	Sub noCache()
		Response.Buffer = True
		Response.Expires = 0
		Response.ExpiresAbsolute = Now() - 1
		Response.CacheControl = "no-cache"
		Response.AddHeader "Expires",Date()
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
	End Sub
	'设置一个Cookies值
	Sub SetCookie(ByVal cooName, ByVal cooValue, ByVal cooCfg)
		Dim n,i,cExp,cDomain,cPath,cSecure
		If isArray(cooCfg) Then
			For i = 0 To Ubound(cooCfg)
				If isDate(cooCfg(i)) Then
					cExp = cDate(cooCfg(i))
				ElseIf Test(cooCfg(i),"int") Then
					If cooCfg(i)<>0 Then cExp = Now()+Int(cooCfg(i))/60/24
				ElseIf Test(cooCfg(i),"domain") or Test(cooCfg(i),"ip") Then
					cDomain = cooCfg(i)
				ElseIf Instr(cooCfg(i),"/")>0 Then
					cPath = cooCfg(i)
				ElseIf cooCfg(i)="True" or cooCfg(i)="False" Then
					cSecure = cooCfg(i)
				End If
			Next
		Else
			If isDate(cooCfg) Then
				cExp = cDate(cooCfg)
			ElseIf Test(cooCfg,"int") Then
				If cooCfg<>0 Then cExp = Now()+Int(cooCfg)/60/24
			ElseIf Test(cooCfg,"domain") or Test(cooCfg,"ip") Then
				cDomain = cooCfg
			ElseIf Instr(cooCfg,"/")>0 Then
				cPath = cooCfg
			ElseIf cooCfg = "True" or cooCfg = "False" Then
				cSecure = cooCfg
			End If
		End If
		If Has(cooValue) Then
			If b_cooen Then
				Use("Aes") : cooValue = Aes.Encode(cooValue)
			End If
		End If
		If Instr(cooName,">")>0 Then
			n = CRight(cooName,">")
			cooName = CLeft(cooName,">")
			Response.Cookies(cooName)(n) = cooValue
		Else
			Response.Cookies(cooName) = cooValue
		End If
		If Has(cExp) Then Response.Cookies(cooName).Expires = cExp
		If Has(cDomain) Then Response.Cookies(cooName).Domain = cDomain
		If Has(cPath) Then Response.Cookies(cooName).Path = cPath
		If Has(cSecure) Then Response.Cookies(cooName).Secure = cSecure
	End Sub
	'获取一个Cookies值
	Function Cookie(ByVal s)
		Dim p,t,coo
		If Instr(s,">") > 0 Then
			p = CLeft(s,">")
			s = CRight(s,">")
		End If
		If Instr(s,":")>0 Then
			t = CRight(s,":")
			s = CLeft(s,":")
		End If
		If Has(p) And Has(s) Then
			If Response.Cookies(p).HasKeys Then
				coo = Request.Cookies(p)(s)
			End If
		ElseIf Has(s) Then
			coo = Request.Cookies(s)
		Else
			Cookie = "" : Exit Function
		End If
		If IsN(coo) Then Cookie = "": Exit Function
		If  b_cooen Then
			Use("Aes") : coo = Aes.Decode(coo)
		End If
		Cookie = Safe(coo,t)
	End Function
	'删除一个Cookies值
	Sub RemoveCookie(ByVal s)
		Dim p,t
		If Instr(s,">") > 0 Then
			p = CLeft(s,">")
			s = CRight(s,">")
		End If
		If Has(p) And Has(s) Then
			If Response.Cookies(p).HasKeys Then
				Response.Cookies(p)(s) = Empty
			End If
		ElseIf Has(s) Then
			Response.Cookies(s) = Empty
			Response.Cookies(s).Expires = Now()
		End If
	End Sub
	'设置缓存记录
	Sub SetApp(ByVal AppName,ByRef AppData)
		Application.Lock
		If IsObject(AppData) Then
			Set Application(AppName) = AppData
		Else
			Application(AppName) = AppData
		End If
		Application.UnLock
	End Sub
	'获取一个缓存记录
	Function GetApp(ByVal AppName)
		If IsN(AppName) Then GetApp = Empty : Exit Function
		If IsObject(Application(AppName)) Then
			Set GetApp = Application(AppName)
		Else
			GetApp = Application(AppName)
		End If
	End Function
	'删除一个缓存记录
	Sub RemoveApp(ByVal AppName)
		Application.Lock
		Application(AppName) = Empty
		Application.UnLock
	End Sub
	'调试ASP变量
	Sub Trace(ByVal o)
		Easp.Ext("Trace")(o)
	End Sub
	'调试ASP变量并显示详细信息
	Sub TraceAll(ByVal o)
		Easp.Ext("Trace").TraceAll(o)
	End Sub
	'获取文本中的图片地址存为一个数组
	Function GetImg(ByVal s)
		GetImg = GetImg__(s,0)
	End Function
	'获取文本中的图像标签存为一个数组
	Function GetImgTag(ByVal s)
		GetImgTag = GetImg__(s,1)
	End Function
	Private Function GetImg__(ByVal s, ByVal t)
		Dim a(), img, Matches, m, i : i = 0
		ReDim a(-1)
		img = "<img([^>]+?)(/?)>"
		If Has(s) And Easp_Test(s, img) Then
			'取消所有的换行和缩进
			s = Replace(s, vbCrLf, " ")
			s = Replace(s, vbTab, " ")
			'正则匹配所有的img标签
			Set Matches = RegMatch(s, "(<img\s[^>]*src\s*=\s*([""|']?))([^""'>]+)(\2[^>]*>)")
			'取出每个img的src存入数组
			For Each m In Matches
				ReDim Preserve a(i)
				a(i) = Easp.IIF(t=0,m.SubMatches(2),m.value)
				i = i + 1
			Next
		End If
		GetImg__ = a
	End Function
	'验证身份证号码
	Private Function isIDCard(ByVal s)
		Dim Ai, BirthDay, arrVerifyCode, Wi, i, AiPlusWi, modValue, strVerifyCode
		isIDCard = False
		If Len(s) <> 15 And Len(s) <> 18 Then Exit Function
		Ai = IIF(Len(s) = 18,Mid(s, 1, 17),Left(s, 6) & "19" & Mid(s, 7, 9))
		If Not IsNumeric(Ai) Then Exit Function
		If Not Test(Left(Ai,6),"^(1[1-5]|2[1-3]|3[1-7]|4[1-6]|5[0-4]|6[1-5]|8[12]|91)\d{2}[01238]\d{1}$") Then Exit Function
		BirthDay = Mid(Ai, 7, 4) & "-" & Mid(Ai, 11, 2) & "-" & Mid(Ai, 13, 2)
		If IsDate(BirthDay) Then
			If cDate(BirthDay) > Date() Or cDate(BirthDay) < cDate("1870-1-1") Then  Exit Function
		Else
			Exit Function
		End If
		arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
		Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
		For i = 0 To 16
			AiPlusWi = AiPlusWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)
		Next
		modValue = AiPlusWi Mod 11
		strVerifyCode = arrVerifyCode(modValue)
		Ai = Ai & strVerifyCode
		If Len(s) = 18 And LCase(s) <> Ai Then Exit Function
		isIDCard = True
	End Function
	'简易的服务端检查表单
	Function CheckForm(ByVal s, ByVal Rule, ByVal Require, ByVal ErrMsg)
		Dim tmpMsg, Msg, i
		tmpMsg = Replace(ErrMsg,"\:",chr(0))
		Msg = IIF(Instr(tmpMsg,":")>0,Split(tmpMsg,":"),Array("有项目不能为空",tmpMsg))
		If Require = 1 And IsN(s) Then
			If Instr(tmpMsg,":")>0 Then
				Alert Replace(Msg(0),chr(0),":") : Exit Function
			Else
				Alert Replace(tmpMsg,chr(0),":") : Exit Function
			End If
		End If
		If Not (Require = 0 And isN(s)) Then
			If Left(Rule,1)=":" Then
				pass = False
				arrRule = Split(Mid(Rule,2),"||")
				For i = 0 To Ubound(arrRule)
					If Test(s,arrRule(i)) Then pass = True : Exit For
				Next
				If Not pass Then Alert(Replace(Msg(1),chr(0),":")) : Exit Function
			Else
				If Not Test(s,Rule) Then Alert(Replace(Msg(1),chr(0),":")) : Exit Function
			End If
		End If
		CheckForm = s
	End Function
	'返回正则验证结果
	Function [Test](ByVal s, ByVal p)
		Dim Pa
		Select Case Lcase(p)
			Case "date"		[Test] = isDate(s) : Exit Function
			Case "idcard"	[Test] = isIDCard(s) : Exit Function
			Case "number"	[Test] = isNumeric(s) : Exit Function
			Case "english"	Pa = "^[A-Za-z]+$"
			Case "chinese"	Pa = "^[\u4e00-\u9fa5]+$"
			Case "username"	Pa = "^[a-zA-Z]\w{2,19}$"
			Case "email"	Pa = "^\w+([-+\.]\w+)*@(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
			Case "int"		Pa = "^[-\+]?\d+$"
			Case "double"	Pa = "^[-\+]?\d+(\.\d+)?$"
			Case "price"	Pa = "^\d+(\.\d+)?$"
			Case "zip"		Pa = "^\d{6}$"
			Case "qq"		Pa = "^[1-9]\d{4,9}$"
			Case "phone"	Pa = "^((\(\+?\d{2,3}\))|(\+?\d{2,3}\-))?(\(0?\d{2,3}\)|0?\d{2,3}-)?[1-9]\d{4,7}(\-\d{1,4})?$"
			Case "mobile"	Pa = "^(\+?\d{2,3})?0?1(3\d|47|5\d|8[056789])\d{8}$"
			Case "url"		Pa = "^(?:(https|http|ftp|rtsp|mms)://(?:([\w!~\*'\(\).&=\+\$%-]+)(?::([\w!~\*'\(\).&=\+\$%-]+))?@)?)?((?:(?:(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d)\.){3}(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d))|(?:(?:(?:[\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+(?:[a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)|localhost))(?::(\d{1,5}))?([#\?/].*)?$"
			Case "domain"	Pa = "^(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
			Case "ip"		Pa = "^((25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)\.){3}(25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)$"
			Case Else Pa = p
		End Select
		[Test] = Easp_Test(CStr(s),Pa)
	End Function
	'正则替换
	Function regReplace(ByVal s, ByVal rule, Byval Result)
		regReplace = Easp_Replace(s,rule,Result,0)
	End Function
	'正则替换多行模式
	Function regReplaceM(ByVal s, ByVal rule, Byval Result)
		regReplaceM = Easp_Replace(s,rule,Result,1)
	End Function
	'正则匹配捕获
	Function regMatch(ByVal s, ByVal rule)
		o_regex.Pattern = rule
		Set regMatch = o_regex.Execute(s)
		o_regex.Pattern = ""
	End Function
	'正则表达式特殊字符转义
	Function RegEncode(ByVal s)
		Dim re,i
		re = Split("\,$,(,),*,+,.,[,?,^,{,|",",")
		For i = 0 To Ubound(re)
			s = Replace(s,re(i),"\"&re(i))
		Next
		RegEncode = s
	End Function
	'替换正则表达式编组
	Function replacePart(ByVal txt, ByVal rule, ByVal part, ByVal replacement)
		If Not Easp_Test(txt, rule) Then
			replacePart = "[not match]"
			Exit Function
		End If
		Dim Match,i,j,ma,pos,uleft,ul
		i = Int(Mid(part,2))-1
		Set Match = RegMatch(txt,rule)(0)
		For j = 0 To Match.SubMatches.Count-1
			ma = Match.SubMatches(j)
			pos = Instr(txt,ma)
			If pos > 0 Then
				ul = Left(txt,pos-1)
				txt = Mid(txt,Len(ul)+1)
				If i = j Then
					replacePart = uleft & ul & Replace(txt,ma,replacement,pos-len(ul),1,0)
					Exit For
				End If
				uleft = uleft & ul & ma
				txt = Mid(txt, Len(ma)+1)
			End If
		Next
		Set Match = Nothing
	End Function
	'检测组件是否安装
	Function isInstall(Byval s)
		On Error Resume Next : Err.Clear()
		isInstall = False
		Dim obj : Set obj = Server.CreateObject(s)
		If Err.Number = 0 Then isInstall = True
		Set obj = Nothing : Err.Clear()
	End Function
	'动态载入文件
	Sub Include(ByVal filePath)
		'On Error Resume Next
		ExecuteGlobal GetIncCode(IncRead(filePath),0)
		If Err.Number<>0 Then
			[error].Msg = " ( " & filePath & " )"
			[error].Raise 1
		End If
		Err.Clear()
	End Sub
	Function getInclude(ByVal filePath)
		'On Error Resume Next
		ExecuteGlobal GetIncCode(IncRead(filePath),1)
		getInclude = EasyAsp_s_html
		If Err.Number<>0 Then
			[error].Msg = " ( " & filePath & " )"
			[error].Raise 1
		End If
		Err.Clear()
	End Function
	'读取文件文本内容
	Function Read(ByVal filePath)
		Dim p, f, o_strm, tmpStr, s_char, t
		s_char = s_charset
		If Instr(filePath,"|")>0 Then
			t = LCase(Trim(CRight(filePath,"|")))
			filePath = CLeft(filePath,"|")
		End If
		If Instr(filePath,">")>0 Then
			s_char = UCase(Trim(CRight(filePath,">")))
			filePath = Trim(CLeft(filePath,">"))
		End If
		p = filePath
		'Easp.WN p
		If Mid(p,2,1)<>":" Then p = Server.MapPath(p)
		If isFile(p) Then
			Set o_strm = Server.CreateObject("ADODB.Stream")
			With o_strm
				.Type = 2
				.Mode = 3
				.Open
				.LoadFromFile p
				.Charset = s_char
				.Position = 2
				tmpStr = .ReadText
				.Close
			End With
			Set o_strm = Nothing
			If s_char = "UTF-8" Then
				Select Case s_bom
					Case "keep"
						'Do Nothing
					Case "remove"
						If Test(tmpStr, "^\uFEFF") Then
							tmpStr = RegReplace(tmpStr, "^\uFEFF", "")
						End If
					Case "add"
						If Not Test(tmpStr, "^\uFEFF") Then
							tmpStr = Chrw(&hFEFF) & tmpStr
						End If
				End Select
			End If
		Else
			tmpStr = ""
			If IsN(t) Then
				tmpStr = "File Not Found: '" & filePath & "'"
			ElseIf t="fso" Then
				[error].Msg = "(" & filePath & ")"
				[error].Raise 2
			End If
		End If
		Read = tmpStr
	End Function
	'读取包含文件内容（无限级）
	Private Function IncRead(ByVal filePath)
		Dim content, rule, inc, incFile, incStr,Match
		content = Read(filePath)
		If isN(content) Then Exit Function
		content = regReplace(content,"<% *?@.*?%"&">","")
		content = regReplace(content,"(<%[^>]+?)(option +?explicit)([^>]*?%"&">)","$1'$2$3")
		rule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
		If Easp_Test(content,rule) Then
			Set inc = regMatch(content,rule)
			For Each Match In inc
				If LCase(Match.SubMatches(0))="virtual" Then
					incFile = Match.SubMatches(1)
				Else
					incFile = Mid(filePath,1,InstrRev(filePath,IIF(Instr(filePath,":")>0,"\","/"))) & Match.SubMatches(1)
				End If
				incStr = IncRead(incFile)
				content = Replace(content,Match,incStr)
			Next
			Set inc = Nothing
		End If
		IncRead = content
	End Function
	'将文本内容转换为ASP代码
	Function GetIncCode(ByVal content, ByVal getHtml)
		Dim tmpStr,code,tmpCode,s_code,st,en
		code = "" : st = 1 : en = Instr(content,"<%") + 2
		s_code = IIF(getHtml=1,"EasyAsp_s_html = EasyAsp_s_html & ","Response.Write ")
		While en > st + 1
			tmpStr = Mid(content,st,en-st-2)
			st = Instr(en,content,"%"&">") + 2
			If Has(tmpStr) Then
				tmpStr = Replace(tmpStr,"""","""""")
				tmpStr = Replace(tmpStr,vbCrLf,"""&vbCrLf&""")
				code = code & s_code & """" & tmpStr & """" & vbCrLf
			End If
			tmpStr = Mid(content,en,st-en-2)
			tmpCode = regReplace(tmpStr,"^\s*=\s*",s_code) & vbCrLf
			If getHtml = 1 Then
				tmpCode = regReplaceM(tmpCode,"response\.write([\( ])", s_code & "$1") & vbCrLf
				tmpCode = regReplaceM(tmpCode,"Easp\.(WC|WN|W)([\( ])", s_code & "$2") & vbCrLf
			End If
			code = code & tmpCode
			en = Instr(st,content,"<%") + 2
		Wend
		tmpStr = Mid(content,st)
		If Has(tmpStr) Then
			tmpStr = Replace(tmpStr,"""","""""")
			tmpStr = Replace(tmpStr,vbcrlf,"""&vbCrLf&""")
			code = code & s_code & """" & tmpStr & """" & vbCrLf
		End If
		If getHtml = 1 Then code = "EasyAsp_s_html = """" " & vbCrLf & code
		GetIncCode = regReplace(code,"(\n\s*\r)+",vbCrLf)
	End Function
	Private Function isFile(ByVal p)
		isFile = False
		If TypeName(o_fso)<>"FileSystemObject" Then Set o_fso = Server.CreateObject(s_fsoName)
		If Mid(p,2,1)<>":" Then p = Server.MapPath(p)
		If o_fso.FileExists(p) Then isFile = True
	End Function
	'加载引用EasyAsp库类
	Sub Use(ByVal f)
		Dim p, o, t : o = f
		p = "easp." & Lcase(o) & ".asp"
		If LCase(o) = "md5" Then o = "o_md5"
		t = Eval("LCase(TypeName(" & o & "))")
		If t = "easyasp_obj" Then
			p = s_path & "core/" & p
			If isFile(p) Then
				Include p
				Execute("Set " & o & " = New EasyAsp_" & f)
			Else
				[error].Msg = "(当前设置 """ & s_path & """ 是错误的)"
				[error].Raise 3
			End If
		End If
	End Sub
	'加载插件
	Public Default Function Ext(ByVal f)
		Dim loaded, p
		f = Lcase(f) : loaded = True
		If Not o_ext.Exists(f) Then
			loaded = False
		Else
			If LCase(TypeName(o_ext(f))) <> "easyasp_" & f Then loaded = False
		End If
		If Not loaded Then
			p = s_plugin & "easp." & f & ".asp"
			If isFile(p) Then
				Include p
				Execute("Set o_ext(""" & f & """) = New EasyAsp_" & f)
			Else
				[error].Msg = "(当前设置 """ & s_path & """ 是错误的)"
				[error].Raise 3
			End If
		End If
		Set Ext = o_ext(f)
	End Function
	'清除加载插件
	Private Sub ClearExt()
		Dim i
		If Has(o_ext) Then
			For Each i In o_ext
				Set o_ext(i) = Nothing
			Next
			o_ext.RemoveAll
		End If
	End Sub
	'Md5加密字符串（32位）
	Function MD5(ByVal s)
		Use("Md5") : MD5 = o_md5(s)
	End Function
	'Md5加密字符串（16位）
	Function MD5_16(ByVal s)
		Use("Md5") : MD5_16 = o_md5.To16(s)
	End Function
	
	'查询CodePage和Charset对照
	Function CharCode(ByVal s)
		Dim co, ch, cn, cnf, i
		co = Array(708,720,28596,1256,1257,852,28592,1250,936,936,950,862,66,874,932,949,1251,1252,1253,1254,1255,1258,20866,21866,28595,28597,28598,38598,50932,51932,52936,65001)
		ch = Array("ASMO-708","DOS-720","iso-8859-6","windows-1256","windows-1257","ibm852","iso-8859-2","windows-1250","gb2312","gbk","big5","DOS-862","cp866","windows-874","shift_jis","ks_c_5601-1987","windows-1251","iso-8859-1","windows-1253","iso-8859-9","windows-1255","windows-1258","koi8-r","koi8-ru","iso-8859-5","iso-8859-7","iso-8859-8","iso-8859-8-i","_autodetect","euc-jp","hz-gb-2312","utf-8")
		cn = Array("阿拉伯字符 (ASMO 708)","阿拉伯字符 (DOS)","阿拉伯字符 (ISO)","阿拉伯字符 (Windows)","波罗的海字符 (Windows)","中欧字符 (DOS)","中欧字符 (ISO)","中欧字符 (Windows)","简体中文 (GB2312)","简体中文 (GBK)","繁体中文 (Big5)","希伯来字符 (DOS)8","西里尔字符 (DOS)","泰语 (Windows)","日语 (Shift-JIS)","朝鲜语","西里尔字符 (Windows)","西欧字符","希腊字符 (Windows)","土耳其字符 (Windows)","希伯来字符 (Windows)","越南字符 (Windows)","西里尔字符 (KOI8-R)","西里尔字符 (KOI8-U)","西里尔字符 (ISO)","希腊字符 (ISO)","希伯来字符 (ISO-Visual)","希伯来字符 (ISO-Logical)","日语 (自动选择)","日语 (EUC)","简体中文 (HZ)","Unicode (UTF-8)")
		If Instr(s,":") Then
			s = CLeft(s,":")
			cnf = True
		End If
		If isNumeric(s) Then
			For i = 0 To Ubound(co)
				If co(i) = s Then
					CharCode = IIF(cnf,cn(i),ch(i))
					Exit Function
				End If
			Next
			CharCode = "utf-8"
		Else
			For i = 0 To Ubound(ch)
				If ch(i) = LCase(s) Then
					CharCode = IIF(cnf,cn(i),co(i))
					Exit Function
				End If
			Next
			CharCode = 65001
		End If
	End Function
	
	Private Function Easp_Test(ByVal s, ByVal p)
		If IsN(s) Then Easp_Test = False : Exit Function
		o_regex.Pattern = p
		Easp_Test = o_regex.Test(CStr(s))
		o_regex.Pattern = ""
	End Function
	Private Function Easp_Replace(ByVal s, ByVal rule, Byval Result, ByVal isM)
		Dim tmpStr,Reg : tmpStr = s
		If Has(s) Then
			If isM = 1 Then o_regex.Multiline = True
			o_regex.Pattern = rule
			tmpStr = o_regex.Replace(tmpStr,Result)
			If isM = 1 Then o_regex.Multiline = False
			o_regex.Pattern = ""
		End If
		Easp_Replace = tmpStr
	End Function
	'取字符串的两头
	Private Function Easp_LR(ByVal s, ByVal m, ByVal t)
		Dim n : n = Instr(s,m)
		If n>0 Then
			If t = 0 Then
				Easp_LR = Left(s,n-1)
			ElseIf t = 1 Then
				Easp_LR = Mid(s,n+Len(m))
			End If
		Else
			Easp_LR = s
		End If
	End Function
End Class
Class EasyAsp_obj : End Class
'Easp内部多参数处理
Private Function Easp_Param(ByVal s)
	Dim arr(1),t : t = Instr(s,":")
	If t > 0 Then
		arr(0) = Left(s,t-1) : arr(1) = Mid(s,t+1)
	Else
		arr(0) = s : arr(1) = ""
	End If
	Easp_Param = arr
End Function
%>
<!--#include file="core/easp.error.asp"-->
<!--#include file="core/easp.db.asp"-->
<!--#include file="core/easp.list.asp"-->
<!--#include file="core/easp.fso.asp"-->
<!--#include file="core/easp.upload.asp"-->
<!--#include file="core/easp.tpl.asp"-->
<!--#include file="core/easp.json.asp"-->
<!--#include file="core/easp.aes.asp"-->
<!--#include file="core/easp.cache.asp"-->
<!--#include file="core/easp.http.asp"-->
<!--#include file="core/easp.xml.asp"-->