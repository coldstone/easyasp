<%
'######################################################################
'## easp.http.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XMLHTTP Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/03/23 23:24:30
'## Description :   Request XMLHttp Data in EasyASP
'## 
'######################################################################
Class EasyAsp_Http
	Public Url, Method, CharSet, Async, User, Password, Html, Headers, Body, Text, SaveRandom
	Public ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
	Private s_data, s_url, s_ohtml, o_rh', a_rh()
	
	Private Sub Class_Initialize
		'编码默认为空，将自动获取编码
		CharSet = ""
		'异步模式关闭
		Async = False
		User = ""
		Password = ""
		s_data = ""
		s_url = ""
		Html = ""
		Headers = ""
		Body = Empty
		Text = Empty
		SaveRandom = True
		'服务器解析超时（毫秒）
		ResolveTimeout = 20000
		'服务器连接超时（毫秒）
		ConnectTimeout = 20000
		'发送数据超时（毫秒）
		SendTimeout = 300000
		'接受数据超时（毫秒）
		ReceiveTimeout = 60000
		Easp.Error(46) = "远程服务器没有响应"
		Easp.Error(47) = "服务器不支持XMLHTTP组件"
		Easp.Error(48) = "要获取的页面地址不能为空"
		Easp.Use "List"
		Set o_rh = Easp.List.New
'		ReDim a_rh(-1)
	End Sub
	
	Private Sub Class_Terminate
		Set o_rh = Nothing
	End Sub

	'建新Easp远程文件操作类实例
	Public Function [New]()
		Set [New] = New EasyAsp_Http
	End Function
	
	'设置要提交的数据
	Public Property Let Data(ByVal s)
		s_data = s
	End Property
	
	'设置请求头信息
	Public Sub SetHeader(ByVal a)
		Dim i,n,v
		If isArray(a) Then
			For i = 0 To Ubound(a)
				n = Replace(Easp.Cleft(a(i),":"),"-","_")
				v = Easp.CRight(a(i),":")
				o_rh(n) = v
			Next
		Else
			n = Replace(Easp.Cleft(a,":"),"-","_")
			v = Easp.CRight(a,":")
			o_rh(n) = v
		End If
	End Sub
	'设置或获取单项请求头信息
	Public Property Let RequestHeader(ByVal n, ByVal v)
		n = Replace(n,"-","_")
		o_rh(n) = v
	End Property
	Public Property Get RequestHeader(ByVal n)
		If Easp.Has(n) Then
			RequestHeader = o_rh(n)
		Else
			RequestHeader = Join(o_rh.Hash,vbCrLf)
		End If
	End Property
'	'传入RequestHeader
	Private Sub SetHeaderTo(ByRef o)
		Dim maps,key
		Set maps = o_rh.Maps
		For Each key In maps
			If Not isNumeric(key) Then
				o.setRequestHeader Replace(key,"_","-"), o_rh(key)
			End If
		Next
		Set maps = Nothing
	End Sub
	
	'属性配置模式下打开连接远程
	Public Function [Open]
		[Open] = GetData(Url, Method, Async, s_data, User, Password)
	End Function
	
	'Get模式取远程页
	Public Function [Get](ByVal uri)
		[Get] = GetData(uri, "GET", Async, s_data, User, Password)
	End Function
	
	'Post模式取远程页
	Public Function Post(ByVal uri)
		Post = GetData(uri, "POST", Async, s_data, User, Password)
	End Function
	
	'获取远程页完整参数模式
	Public Function GetData(ByVal uri, ByVal m, ByVal async, ByVal data, ByVal u, ByVal p)
		Dim o,chru
		'建立XMLHttp对象
'		If Easp.isInstall("MSXML2.serverXMLHTTP") Then
'			Set o = Server.CreateObject("MSXML2.serverXMLHTTP")
'		ElseIf Easp.isInstall("MSXML2.XMLHTTP") Then
			Set o = Server.CreateObject("MSXML2.XMLHTTP")
'		ElseIf Easp.isInstall("Microsoft.XMLHTTP") Then
'			Set o = Server.CreateObject("Microsoft.XMLHTTP")
'		Else
'			Easp.Error.Raise 47
'			Exit Function
'		End If
		'设置超时时间
		'o.SetTimeOuts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
		'抓取地址
		If Easp.IsN(uri) Then Easp.Error.Raise 48 : Exit Function
		'通过URL临时指定编码
		If Easp.Test(uri,"^[\w\d-]+>https?://") Then
			CharSet = Easp.CLeft(uri,">")
			uri = Easp.CRight(uri,">")
		End If
		s_url = uri
		'方法：POST或GET
		m = Easp.IIF(Easp.Has(m),UCase(m),"GET")
		'异步
		If Easp.IsN(async) Then async = False
		'构造Get传数据的URL
		If m = "GET" And Easp.Has(data) Then uri = uri & Easp.IIF(Instr(uri,"?")>0, "&", "?") & Serialize__(data)
		'打开远程页
		If Easp.Has(u) Then
			'如果有用户名和密码
			o.open m, uri, async, u, p
		Else
			'匿名
			o.open m, uri, async
		End If
		If m = "POST" Then
			If Not o_rh.HasIndex("Content_Type") Then
				o_rh("Content_Type") = "application/x-www-form-urlencoded"
			End If
			SetHeaderTo o
			'有发送的数据
			o.send Serialize__(data)
		Else
			SetHeaderTo o
			o.send
		End If
		'检测返回数据
		If o.readyState <> 4 Then
			GetData = "error:server is down"
			Set o = Nothing
			Easp.Error.Raise 46
			Exit Function
		ElseIf o.Status = 200 Then
			Headers = o.getAllResponseHeaders()
			Body = o.responseBody
			If Easp.IsN(CharSet) Then
			  Text = o.responseText
				'从Header中提取编码信息
				If Easp.Test(Headers,"charset=([\w-]+)") Then
					CharSet = Easp.RegReplace(Headers,"([\s\S]+)charset=([\w-]+)([\s\S]+)","$2")
				'如果是Xml文档，从文档中提取编码信息
				ElseIf Easp.Test(Headers,"Content-Type: ?text/xml") Then
					CharSet = Easp.RegReplace(Text,"^<\?xml\s+[^>]+encoding\s*=\s*""([^""]+)""[^>]*\?>([\s\S]+)","$1")
				'从文件源码中提取编码
				ElseIf Easp.Test(Text,"<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>") Then
					CharSet = Easp.RegReplace(Text,"([\s\S]+)<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>([\s\S]+)","$2")
				End If
				'Easp.WNH CharSet
				'如果无法获取远程页的编码则继承Easp的编码设置
				If Easp.IsN(CharSet) Then CharSet = Easp.CharSet
			End If
			GetData = Bytes2Bstr__(Body, CharSet)
		Else
			GetData = "error:" & o.Status & " " & o.StatusText
		End If
		Set o = Nothing
		s_ohtml = GetData
		Html = s_ohtml
	End Function
		
	'按正则查找符合的第一个字符串
	Public Function Find(ByVal rule)
		Find = Find_(s_ohtml, rule)
	End Function
	Public Function Find_(ByVal s, ByVal rule)
		If Easp.Test(s,rule) Then Find_ = Easp.RegReplace(s,"([\s\S]*)("&rule&")([\s\S]*)","$2")
	End Function
	
	'按正则查找符合的第一个字符串，可按正则编组选择其中的一部分
	Public Function [Select](ByVal rule, ByVal part)
		[Select] = Select_(s_ohtml, rule, part)
	End Function
	Public Function Select_(ByVal s, ByVal rule, ByVal part)
		If Easp.Test(s,rule) Then
			'$0匹配字符串本身
			part = Replace(part,"$0",Find_(s,rule))
			'按正则编组分别替换
			Select_ = Easp.RegReplace(s,"(?:[\s\S]*)(?:"&rule&")(?:[\s\S]*)",part)
		End If
	End Function
	
	'按正则查找符合的字符串组，返回数组
	Public Function Search(ByVal rule)
		Search = Search_(s_ohtml, rule)
	End Function
	Public Function Search_(ByVal s, ByVal rule)
		Dim matches,match,arr(),i : i = 0
		Set matches = Easp.RegMatch(s,rule)
		ReDim arr(matches.Count-1)
		For Each match In matches
			arr(i) = match.Value
			i = i + 1
		Next
		Set matches = Nothing
		Search_ = arr
	End Function
	
	'按标签查找字符串
	Public Function SubStr(ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
	'tagStart - 要截取的部分的开头
	'tagEnd   - 要截取的部分的结尾
	'tagSelf  - 结果是否包括tagStart和tagEnd
	'           (0或空:不包括,1:包括,2:只包括tagStart,3:只包括tagEnd)
		SubStr = SubStr_(s_ohtml,tagStart,tagEnd,tagSelf)
	End Function
	Public Function SubStr_(ByVal s, ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
		Dim posA, posB, first, between
		posA = instr(1,s,tagStart,1)
		If posA=0 Then SubStr_ = "源代码中不包括此开始标签" : Exit Function
		posB = instr(PosA+Len(tagStart),s,tagEnd,1) 
		If posB=0 Then SubStr_ = "源代码中不包括此结束标签" : Exit Function
		Select Case tagSelf
			Case 1
				first = posA
				between = posB+len(tagEnd)-first
			Case 2
				first = posA
				between = posB-first
			Case 3
				first = posA+len(tagStart)
				between = posB+len(tagEnd)-first
			Case Else
				first = posA+len(tagStart)
				between = posB-first
		End Select
		SubStr_ = Mid(s,first,between)
	End Function
	
	'保存远程图片到本地
	Public Function SaveImgTo(ByVal p)
		SaveImgTo = SaveImgTo_(s_ohtml,p)
	End Function
	Public Function SaveImgTo_(ByVal s, ByVal p)
		Dim a,b, i, img, ht, tmp, src
		'取得图片地址
		a = Easp.GetImg(s)
		b = Easp.GetImgTag(s)
		If Easp.Has(a) Then
			Easp.Use "Fso"
			For i = 0 To Ubound(a)
				If SaveRandom Then
					img = Easp.DateTime(Now,"ymmddhhiiss"&Easp.RandStr("5:0123456789")) & Mid(a(i),InstrRev(a(i),"."))
				Else
					img = Mid(a(i),InstrRev(a(i),"/")+1)
				End If
				Set ht = Easp.Http.New
				ht.Get "UTF-8>" & TransPath(s_url, a(i))
				tmp = Easp.Fso.SaveAs(p & img, ht.Body)
				Set ht = Nothing
				If tmp Then
					'Easp.WN "b(i)=> " & Easp.HtmlEncode(b(i))
					src = Easp.RegReplace(b(i),"(<img\s[^>]*src\s*=\s*([""|']?))("&a(i)&")(\2[^>]*>)","$1"&p&img&"$4")
					'Easp.WN "src=> " & Easp.HtmlEncode(src)
					s = Replace(s,b(i),src)
				End If
			Next
		End If
		SaveImgTo_ = s
	End Function
	
	'启用Ajax代理
	Public Sub AjaxAgent()
		Easp.NoCache()
		Dim u, qs, qskey, qf, qfkey, m
		'取得目标地址
		u = Easp.Get("easpurl")
		If Easp.IsN(u) Then Easp.WE "error:Invalid URL"
		If Instr(u,"?")>0 Then
			qs = "&" & Easp.CRight(u,"?")
			u = Easp.CLeft(u,"?")
		End If
		'传url参数
		If Request.QueryString()<>"" Then
			For Each qskey In Request.QueryString
				If qskey<>"easpurl" Then qs = qs & "&" & qskey & "=" & Request.QueryString(qskey)
			Next
		End If
		u = u & Easp.IfThen(Easp.Has(qs),"?" & Mid(qs,2))
		'Easp.WC u
		'如果是Post则同时传Form数据
		m = Request.ServerVariables("REQUEST_METHOD")
		If m = "POST" Then
			If Request.Form()<>"" Then
				For Each qfkey In Request.Form
					qf = qf & "&" & qfkey & "=" & Request.Form(qfkey)
				Next
				Data = Mid(qf,2)
			End If
			Easp.WE Post(u)
		Else
			Easp.WE [Get](u)
		End If
	End Sub
	
	'转换绝对路径
	Function TransPath(ByVal u, ByVal p)
		'如果本来就是绝对路径则直接取出
		If Left(p,7)="http://" Or Left(p,8)="https://" Then TransPath = p : Exit Function
		Dim tmp,ser, fol
		'页面地址
		tmp = Easp.CLeft(u,"?")
		'服务器地址
		If Left(u,7)<>"http://" And Left(u,8)<>"https://" Then
			ser = ""
		Else
			ser = Easp.RegReplace(tmp,"^(https?://[a-zA-Z0-9-.]+)/(.+)$","$1")
		End If
		'页面所在路径
		fol = Mid(tmp,1,InstrRev(tmp,"/"))
		TransPath = Easp.IIF(Left(p,1) = "/", ser, fol) & p
	End Function
	
	'url参数化
	Private Function Serialize__(ByVal a)
		Dim tmp, i, n, v : tmp = ""
		If Easp.IsN(a) Then Exit Function
		If isArray(a) Then
			For i = 0 To Ubound(a)
				n = Easp.CLeft(a(i),":")
				v = Easp.CRight(a(i),":")
				tmp = tmp & "&" & n & "=" & Server.URLEncode(v)
			Next
			If Len(tmp)>1 Then tmp = Mid(tmp,2)
			Serialize__ = tmp
		Else
			Serialize__ = a
		End If
	End Function
	
	'编码转换
	Private Function Bytes2Bstr__(ByVal s, ByVal char) 
		dim oStrm
		set oStrm = Server.CreateObject("Adodb.Stream")
		With oStrm
			.Type = 1
			.Mode =3
			.Open
			.Write s
			.Position = 0
			.Type = 2
			.Charset = CharSet
			Bytes2Bstr__ = .ReadText
			.Close
		End With
		set oStrm = nothing
	End Function
End Class
%>