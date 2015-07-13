<%
'######################################################################
'## easp.error.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Exception Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/11/13 22:01:30
'## Description :   Deal with the EasyAsp Exception
'##
'######################################################################
Class EasyAsp_Error
	Private b_debug, b_redirect
	Private i_errNum, i_delay
	Private s_errStr, s_title, s_url, s_css, s_msg
	Private o_err
	Private Sub Class_Initialize
		i_errNum    = ""
		i_delay     = 3000
		s_title     = "发生错误啦"
		b_debug     = Easp.Debug
		b_redirect  = True
		s_url       = "javascript:history.go(-1)"
		Set o_err   = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub
	'是否开启调试状态（开启后返回开发者错误信息）
	Public Property Get [Debug]
		[Debug] = b_debug
	End Property
	Public Property Let [Debug](ByVal b)
		b_debug = b
	End Property
	'设置或读取自定义的错误代码和错误信息
	Public Default Property Get E(ByVal n)
		If IsNumeric(n) Then n = CLng(n)
		If o_err.Exists(n) Then
			E = o_err(n)
		Else
			E = "未知错误"
		End If
	End Property
	Public Property Let E(ByVal n, ByVal s)
		If Easp.Has(n) And Easp.Has(s) Then
			If n > "" Then
				If IsNumeric(n) Then n = CLng(n)
				o_err(n) = s
			End If
		End If
	End Property
	'取最后一次发生错误的代码
	Public Property Get LastError
		LastError = i_errNum
	End Property
	'设置和读取错误信息标题
	Public Property Get Title
		Title = s_title
	End Property
	Public Property Let Title(ByVal s)
		s_title = s
	End Property
	'设置和读取自定义的附加错误信息
	Public Property Get Msg
		Msg = s_msg
	End Property
	Public Property Let Msg(ByVal s)
		s_msg = s
	End Property
	'设置和读取页面是否自动转向
	Public Property Get [Redirect]
		[Redirect] = b_redirect
	End Property
	Public Property Let [Redirect](ByVal b)
		b_redirect = b
	End Property
	'设置和读取发生错误后的跳转页地址
	Public Property Get Url
		Url = s_url
	End Property
	Public Property Let Url(ByVal s)
		s_url = s
	End Property
	'设置和读取自动跳转页面等待时间（秒）
	Public Property Get Delay
		Delay = i_delay / 1000
	End Property
	Public Property Let Delay(ByVal i)
		i_delay = i * 1000
	End Property
	'设置和读取显示错误信息DIV的CSS样式名称
	Public Property Get ClassName
		ClassName = s_css
	End Property
	Public Property Let ClassName(ByVal s)
		s_css = s
	End Property
	'生成一个错误
	Public Sub Raise(ByVal n)
		If Easp.isN(n) Then Exit Sub
		i_errNum = n
		If b_debug Then
			Easp.WE ShowMsg(o_err(n) & s_msg, 1)
		End If
		s_msg = ""
	End Sub
	'立即抛出一个错误信息
	Public Sub Throw(ByVal msg)
		If Left(msg,1) = ":" Then
			msg = Mid(msg,2)
			If isNumeric(msg) Then msg = CLng(msg)
			If o_err.Exists(msg) Then msg = o_err(msg)
		End If
		Easp.W ShowMsg(msg,0)
	End Sub
	'显示已定义的所有错误代码及信息
	Public Sub Defined()
		Dim key
		If Easp.Has(o_err) Then
			For Each key In o_err
				Easp.Wn key & " : " & o_err(key)
			Next
		End If
	End Sub
	'显示错误信息框
	Private Function ShowMsg(ByVal msg, ByVal t)
		Dim s,x
		s = "<fieldset id=""easpError""" & Easp.IfThen(Easp.Has(s_css)," class=""" & s_css & """") & ">" & vbCrLf
		s = s & "	<legend>" & s_title & "</legend>" & vbCrLf
		s = s & "	<p class=""msg"">" & msg & "</p>" & vbCrLf
		x = Easp.IIF(s_url = "javascript:history.go(-1)", "返回", "继续")
		If t = 1 Then
			If Err.Number<>0 Then
				s = s & "	<ul class=""dev"">" & vbCrLf
				s = s & "		<li class=""info"">以下信息针对开发者s：</li>" & vbCrLf
				s = s & "		<li>错误代码：0x" & Hex(Err.Number) & "</li>" & vbCrLf
				s = s & "		<li>错误描述：" & Err.Description & "</li>" & vbCrLf
				s = s & "		<li>错误来源：" & Err.Source & "</li>" & vbCrLf
				s = s & "	</ul>" & vbCrLf
			End If
		Else
			If b_redirect Then
				s = s & "	<p class=""back"">页面将在" & i_delay/1000 & "秒钟后跳转，如果浏览器没有正常跳转，<a href=""" & s_url & """>请点击此处" & x & "</a></p>" & vbCrLf
				s_url = Easp.IIF(Left(s_url,11) = "javascript:", Mid(s_url,12), "location.href='" & s_url & "';")
				s = s & Easp.JsCode("setTimeout(function(){" & s_url & "}," & i_delay & ");")
			Else
				s = s & "	<p class=""back""><a href=""" & s_url & """>请点击此处" & x & "</a></p>" & vbCrLf
			End If
		End If
		s = s & "</fieldset>" & vbCrLf
		ShowMsg = s
	End Function
End Class
%>