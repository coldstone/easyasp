<%
'#################################################################################
'##	easp.pyo.asp
'##	------------------------------------------------------------------------------
'##	Feature		:	PinYin Online Transfer for EasyAsp
'##	Version		:	v1.0
'##	Author		:	Coldstone(coldstone[at]qq.com)
'##	Update Date	:	2010/12/16 15:56:57
'## Special Thanks	: kdd.cc
'##	Description	:	这是一个将汉字转换为拼音的插件，支持将UTF-8或GBK编码下的简、繁汉字甚至是生僻
'##               字转换为汉语拼音，支持多音字的识别，可以输出六种格式的汉语拼音。由于本插件采用
'##               在线转换，所以需要服务器能访问互联网。
'## 使用说明：
'##    1、基本用法： Easp.Ext("pyo")("长春市长")
'##                结果：cháng chūn shì zhǎng
'##    2、本插件可以返回6种格式的拼音，用 Type 属性设置，设置方法如下：
'##       Easp.Ext("pyo").Type = <number>
'##       <number> 的值可以是以下数字
'##        1 - 带声调的汉语拼音：cháng chūn shì zhǎng
'##        2 - 首字母大写：Cháng Chūn Shì Zhǎng
'##        3 - 不带声调的拼音：chang chun shi zhang
'##        4 - 声调用数字表示的拼音：chang2 chun1 shi4 zhang3
'##        5 - 汉字注音：ㄔㄤˊ ㄔㄨㄣ ㄕㄧˋ ㄓㄤˇ
'##        6 - 拼音首字母：ccsz
'##    3、也可以用下面的方法而不用设置此属性：
'##       Easp.Ext("pyo").PY("长春市长",6)  '结果：ccsz
'##    4、可以用Space属性设置是否保留每个拼音之间的空格：
'##       Easp.Ext("pyo").Space = <True|False>
'##
'## 特别说明：此插件发布时编码为utf-8，如果要使用在gbk编码下，请自行转换此文档的编码。
'#################################################################################
Class EasyAsp_Pyo

	Private s_author, s_version
	Private i_type
	Private b_space
	Private o_http

	Private Sub Class_Initialize()
		s_author	= "coldstone"
		s_version	= "1.0"
		Easp.Use "Http"
		Set o_http = Easp.Http.New
		i_type = 1
		b_space = True
	End Sub
	Private Sub Class_Terminate()
		Set o_http = Nothing
	End Sub

	Public Property Get Author()
		Author = s_author
	End Property
	Public Property Get Version()
		Version = s_version
	End Property
	Public Property Let [Type](ByVal n)
		i_type = n
	End Property
	Public Property Let [Space](ByVal b)
		b_space = b
	End Property
	
	Public Default Function PinYin(ByVal s)
		PinYin = PY(s, i_type)
	End Function
	
	Public Function PY(ByVal s, ByVal t)
		If Easp.IsN(s) Then Exit Function
		Dim p,u,tmp,arr,i
		u = "http://py.kdd.cc/Unicode/index.asp"
		o_http.Data = "u=" & Easp.IIF(t=6,3,t) & "&wz=" & s
		o_http.SetHeader "referer:" & u
		o_http.Post u
		p = o_http.SubStr("<div name=v  id=v>","<br><br>",0)
		p = Easp.RegReplace(p,"</?font[^>]*>","")
		arr = Split(p,"<br>")
		If Len(s) = 1 Then
			tmp = Easp.CLeft(arr(0)," ")
			If t = 6 then tmp = Left(tmp,1)
		Else
			For i = 0 To Ubound(arr)
				If i Mod 2 = 0 Then tmp = tmp & arr(i)
			Next
		End If
		If t = 6 Then
			arr = Split(tmp," ")
			tmp = ""
			For i = 0 To Ubound(arr)
				tmp = tmp & Left(arr(i),1)
			Next
		End If
		If Not b_space Then tmp = Replace(tmp," ", "")
		PY = tmp
	End Function

End Class
%>