<%
'#################################################################################
'## easp.trace.asp
'## ------------------------------------------------------------------------------
'##  Feature      : EasyASP Variable Tracing Plugin
'##  Version      : v1.3
'##  For EasyASP  :  3.0+
'##  Author       : Coldstone(coldstone[at]qq.com)
'##  Update Date  : 2014-04-26 11:13:19
'##  Description  : 
'##       此插件用于测试ASP的各类变量，变量可以是字符串、数组、二维数组、记录集、Dictionary对象、
'##       EasyASP List对象、Connection对象等，使用方法如下：
'##         Easp.Ext("Trace")(variable)
'##       同时也可以用此方法输出当前的环境变量信息，使用方法如下：
'##         Easp.Ext("Trace")(":get")    - 输出Request.QueryString变量
'##         Easp.Ext("Trace")(":post")   - 输出Request.Form变量
'##         Easp.Ext("Trace")(":cookie") - 输出Request.Cookies变量
'##         Easp.Ext("Trace")(":server") - 输出Request.ServerVariables变量
'##         Easp.Ext("Trace")(":session")- 输出当前Session
'##         Easp.Ext("Trace")(":app")    - 输出已缓存的Application
'##       使用本插件还可以查看数据库结构信息，使用方法如下：
'##         Easp.Ext("Trace")(":db")     - 查看当前数据库的数据表和视图信息
'##         Easp.Ext("Trace")(":db.表名") - 查看某一数据表的详细信息
'##       部分项目如要查看详细数据信息，还可使用：
'##         Easp.Ext("Trace").TraceAll(variable)
'##       特别感谢：Jorkin提供Trace函数原型。
'#################################################################################
Class EasyASP_Trace

	Private s_author, s_version, tpl

	Private Sub Class_Initialize()
		s_author	= "coldstone"
		s_version	= "1.0"
		Set tpl = Easp.Tpl.New
		tpl.TagMask = "{{*}}"
		tpl.LoadStr "<style>.easp-trace{width:90%;font-size:12px;font-family:Consolas;margin:10px auto;padding:0;background-color:#FFF;}.easp-trace h3,.easp-trace h4{font-size:12px;margin:0;line-height:24px;text-align:center;background-color:#999;border:1px solid #555;color:#FFF;border-bottom:none;}.easp-trace h4{padding:5px;line-height:1.5em;text-align:left;background-color:#FFC;color:#000; font-weight:normal;}.easp-trace h4 strong{color:red;}.easp-trace table{width:100%;margin:0;padding:0;border-collapse:collapse;border:1px solid #555;border-bottom:none;}.easp-trace th{background-color:#EEE;white-space:nowrap;}.easp-trace thead th{background-color:#CCC;}.easp-trace th,.easp-trace td{font-size:12px;border:1px solid #999;padding:4px;word-break:break-all;}.easp-trace span.info{color:#F30;}</style><div class=""easp-trace""><h3>EasyASP变量调试</h3><h4>此次调试变量的类型是 <strong>{{type}}</strong> ，{{#if @count!=''}}共有 <strong>{{count}}</strong> 条数据，{{/#if}}以下是其中的{{#if @top>0}}前 <strong>{{top}}</strong> 条，如果要查看全部数据，请使用Easp.Ext(""Trace"").TraceAll方法{{#else}}数据{{/#if}}：</h4>{{table}}</div>"
		tpl "top", 0
	End Sub
	Private Sub Class_Terminate()
		Set tpl = Nothing
	End Sub
	'Set Property
	Public Property Get Author()
		Author = s_author
	End Property
	Public Property Get Version()
		Version = s_version
	End Property
	Public Function [New]
		Set [New] = New EasyASP_Trace
	End Function
	'测试变量
	Public Default Sub Trace(ByVal o)
		Dim t : Set t = [New]
		t.Tracing o,0
		t.Show()
		Set t = Nothing
	End Sub
	Public Sub TraceAll(ByVal o)
		Dim t : Set t = [New]
		t.Tracing o,1
		t.Show()
		Set t = Nothing
	End Sub
	Public Sub Show()
		Easp.Print tpl.GetHtml()
	End Sub
	Private Function GetTable(ByVal n)
		Select Case n
			Case 0 : GetTable = "<table><thead><tr><th width=""20%"">{{cname}}</th><th width=""80%"">{{cvalue}}</th></tr></thead>{{#:loop}}<tr><th>{{name}}</th><td>{{value}}</td></tr>{{/#:loop}}</table>"
			Case 1 : GetTable = "<table><thead><tr><th width=""5%"">{{cno}}</th><th width=""15%"">{{cname}}</th><th width=""80%"">{{cvalue}}</th></tr></thead>{{#:loop}}<tr><th>{{no}}</th><th>{{name}}</th><td>{{value}}</td></tr>{{/#:loop}}</table>"
			Case 2 : GetTable = "<table><thead><tr><th width=""3%"">{{cno}}</th>{{#:col}}<th>{{field}}</th>{{/#:col}}</tr></thead>{{#:rs}}<tr><th>{{i}}</th>{{#:fields}}<td>{{value}}</td>{{/#:fields}}</tr>{{/#:rs}}</table>"
			Case 3 : GetTable = "<table><thead><tr><th width=""20%"">{{cname}}</th><th width=""80%"">{{cvalue}}</th></tr></thead></table>{{#:rs}}<h4>第<strong> {{i}} </strong>条数据：</h4><table>{{#:loop}}<tr><th width=""20%"">{{name}}</th><td width=""80%"">{{value}}</td></tr>{{/#:loop}}</table>{{/#:rs}}"
			Case 4 : GetTable = "<table><tr><td>{{value}}</td></tr></table>"
			Case 5 : GetTable = "<h4>{{info}}</h4>"
			Case 6 : GetTable = "{{#:table}}<h4><strong>{{tableorview}}：{{name}}</strong></h4><table><thead><tr><th width=""20%"">字段名</th><th width=""20%"">字段类型/大小</th><th width=""10%"">允许空</th><th width=""10%"">默认值</th><th width=""40%"">说明</th></tr></thead>{{#:loop}}<tr><th>{{field}}</th><td style=""text-align:center;"">{{datatype}}</td><td style=""text-align:center;"">{{nullable}}</td><td style=""text-align:center;"">{{default}}</td><td>{{desc}}</td></tr>{{/#:loop}}</table>{{/#:table}}"
		End Select
	End Function

	'调试函数
	Public Sub Tracing(ByVal o, ByVal t)
		Dim s,i,j,dbtmp : i = 0 : j = 0
		Select Case VarType(o)
			Case vbEmpty
				tpl.TagStr "table", GetTable(4)
				tpl "type", "空值"
				tpl "value", "[Empty]"
			Case vbNull
				tpl.TagStr "table", GetTable(4)
				tpl "type", "Null值"
				tpl "value", "[Null]"
			Case vbString
				If TypeName(o) = "Connection" Then
						TraceDb o, t
						tpl "type", "Connection对象 "
						tpl "info", "您可以用 Easp.Ext(""Trace"").TraceAll 方法查看该连接对象全部表和视图的详细信息。"
						If o.State = 0 Then o.Open
				ElseIf o="" Then
					tpl.TagStr "table", GetTable(4)
					tpl "type", "空字符串"
					tpl "value", "[Empty String]"
				Else
					Select Case Lcase(o)
						Case ":cookie", "cookies"
							tpl "type", "Cookies集合"
							If Request.Cookies.Count = 0 Then
								tpl.TagStr "table", GetTable(4)
								tpl "value", "您的电脑上没有任何本站的Cookies数据"
							Else
								tpl.TagStr "table", GetTable(1)
								tpl "cno", "序号"
								tpl "cname", "名称"
								tpl "cvalue", "值"
								For Each i In Request.Cookies
									If Request.Cookies(i).HasKeys Then
										For Each j In Request.Cookies(i)
											tpl "no", s+1
											tpl "name", "cookies("""&i&""")("""&j&""")"
											tpl "value", Easp.Str.HtmlEncode(Easp.Cookie(i&">"&j))
											tpl.Update "loop"
											s = s + 1
										Next
									Else
										tpl "no", s+1
										tpl "name", "Cookies("""&i&""")"
										tpl "value", Easp.Str.HtmlEncode(Easp.Cookie(i))
										tpl.Update "loop"
										s = s + 1
									End If
								Next
								tpl "count", s
							End If
						Case ":get", ":querystring"
							tpl "type", "Request.QueryString集合"
							If Request.QueryString.Count = 0 Then
								tpl.TagStr "table", GetTable(4)
								tpl "value", "没有任何QueryString参数被传递"
							Else
								tpl.TagStr "table", GetTable(0)
								tpl "cname", "参数名称"
								tpl "cvalue", "参数值"
								tpl "count", Request.QueryString.Count
								For Each i In Request.QueryString
									tpl "name", "QueryString("""&i&""")"
									tpl "value", Easp.Str.HtmlEncode(Request.QueryString(i))
									tpl.Update "loop"
								Next
							End If
						Case ":post", ":form"
							tpl "type", "Request.Form集合"
							If Request.Form.Count = 0 Then
								tpl.TagStr "table", GetTable(4)
								tpl "value", "没有任何表单数据被提交"
							Else
								tpl.TagStr "table", GetTable(0)
								tpl "cname", "表单项名称"
								tpl "cvalue", "提交的值"
								tpl "count", Request.Form.Count
								For Each i In Request.Form
									tpl "name", "Form("""&i&""")"
									tpl "value", Easp.Str.HtmlEncode(Request.Form(i))
									tpl.Update "loop"
								Next
							End If
						Case ":server", ":servervariables"
							tpl "type", "ServerVariables变量"
							tpl.TagStr "table", GetTable(0)
							tpl "cname", "名称"
							tpl "cvalue", "值"
							tpl "count", Request.ServerVariables.Count
							For Each i In Request.ServerVariables
								tpl "name", i
								tpl "value", Easp.Str.HtmlEncode(Request.ServerVariables(i))
								tpl.Update "loop"
							Next
						Case ":app", ":application"
							tpl "type", "Application缓存"
							If Application.Contents.Count=0 Then
								tpl.TagStr "table", GetTable(4)
								tpl "value", "目前没有任何缓存"
							Else
								tpl.TagStr "table", GetTable(0)
								tpl "cname", "缓存名称"
								tpl "cvalue", "缓存值"
								tpl "count", Application.Contents.Count
								For Each i In Application.Contents
									tpl "name", i
									ShowValue Application(i), "Application", i
									tpl.Update "loop"
								Next
							End If
						Case ":session"
							tpl "type", "Session对象"
							tpl.TagStr "table", GetTable(0)
							tpl "cname", "Session名称"
							tpl "cvalue", "Session值"
							tpl "count", Session.Contents.Count
							On Error Resume Next
							UpdateLoop "Session.CodePage"
							UpdateLoop "Session.LCID"
							UpdateLoop "Session.SessionID"
							UpdateLoop "Session.Timeout"
							On Error Goto 0
							For Each i In Session.Contents
								tpl "name", "Session("""&i&""")"
								ShowValue Session(i), "Session", i
								tpl.Update "loop"
							Next
						Case ":db"
						  Set dbtmp = Easp.db.GetConn()
							TraceDb dbtmp, t
							If dbtmp.State = 0 Then dbtmp.Open
						Case Else
							If Easp.Str.Test(o,"^:db\.(.+)$") Then
						    Set dbtmp = Easp.db.GetConn()
								s = Easp.Str.Replace(o,"^:db\.(.+)$","$1")
								tpl "type", "数据表"
								tpl.TagStr "table", GetTable(6)
								TraceTable s, dbtmp
								If dbtmp.State = 0 Then dbtmp.Open
							Else
								tpl.TagStr "table", GetTable(4)
								tpl "type", "字符串"
								tpl "Value", Easp.Str.HtmlEncode(o)
							End If
					End Select
				End If
			Case vbObject
				Select Case TypeName(o)
					Case "Nothing","Empty"
						tpl.TagStr "table", GetTable(4)
						tpl "type", "空对象"
						tpl "value", "[Empty Object]"
					Case "Recordset"
						tpl "type", "记录集"
						If o.State = 0 Then
							tpl.TagStr "table", GetTable(4)
							tpl "value", "此记录集对象已关闭"
						Else
							If Easp.IsN(o) Then
								tpl.TagStr "table", GetTable(4)
								tpl "value", "此记录集对象为空记录集，没有数据"
							Else
								On Error Resume Next
								Set o = o.Clone
								On Error Goto 0
								If o.RecordCount = 1 Then
									tpl.TagStr "table", GetTable(0)
									tpl "type", "单条记录集"
									tpl "cname", "字段名"
									tpl "cvalue", "字段值"
									For j = 0 To o.Fields.Count-1
										tpl "name", o.Fields(j).Name
										tpl "value", Easp.Str.HtmlEncode(o.Fields(j).Value)
										tpl.Update "loop"
									Next
								Else
									tpl "count", o.RecordCount
									If t = 0 Then
										tpl.TagStr "table", GetTable(2)
										tpl "cno", "序号"
										tpl.Tag("top") = Easp.IIF(o.RecordCount>30, 30, 0)
										For j = 0 To o.Fields.Count-1
											tpl "field", o.Fields(j).Name
											tpl.Update "col"
										Next
										o.MoveFirst
										While i<30 And Not o.Eof
											tpl "i", i+1
											For j = 0 To o.Fields.Count-1
												tpl "value", Easp.Str.HtmlEncode(o.Fields(j).value)
												tpl.Update "fields"
											Next
											tpl.Update "rs"
											i = i + 1
											o.MoveNext
										Wend
									ElseIf t = 1 Then
										tpl.TagStr "table", GetTable(3)
										tpl "cname", "字段名"
										tpl "cvalue", "字段值"
										o.MoveFirst
										While Not o.Eof
											tpl "i", i+1
											For j = 0 To o.Fields.Count-1
												tpl "name", o.Fields(j).Name
												tpl "value", Easp.Str.HtmlEncode(o.Fields(j).value)
												tpl.Update "loop"
											Next
											tpl.Update "rs"
											i = i + 1
											o.MoveNext
										Wend
									End If
								End If
							End If
						End If
					Case "Dictionary"
						tpl "type", "Dictionary对象"
						If o.Count = 0 Then
							tpl.TagStr "table", GetTable(4)
							tpl "value", "此Dictionary对象是空的，还没有任何键值"
						Else
							tpl.TagStr "table", GetTable(0)
							tpl "cname", "键名"
							tpl "cvalue", "键值"
							tpl "count", o.Count
							tpl.Tag("top") = Easp.IIF(o.Count>50 And t=0,50,0)
							For Each i In o
								If t = 0 And j>=50 Then Exit For
								tpl "name", i
								'tpl "value", Easp.Str.HtmlEncode(o(i))
								ShowValue o(i), "", 0
								tpl.Update "loop"
								j = j + 1
							Next
						End If
					Case "EasyASP_List"
						tpl "type", "Easp数组对象(List)"
						If o.Size = 0 Then
							tpl.TagStr "table", GetTable(4)
							tpl "value", "此List对象是空的，还没有任何元素"
						Else
							tpl.TagStr "table", GetTable(1)
							tpl "cno", "下标"
							tpl "cname", "键名"
							tpl "cvalue", "键值"
							tpl "count", o.Size
							tpl.Tag("top") = Easp.IIF(o.Size>50 And t=0,50,0)
							For i = 0 To o.End
								If t = 0 And j>=50 Then Exit For
								tpl "no", i
								tpl "name", o.IndexHash(i)
								'tpl "value", Easp.Str.HtmlEncode(o(i))
								ShowValue o(i), "", 0
								tpl.Update "loop"
								j = j + 1
							Next
						End If
				End Select
			Case vbArray,8194,8204,8209
				Dim arrType, size1, size2
				On Error Resume Next
				size1 = Ubound(o)
				size2 = Ubound(o,2)
				arrType = Easp.IIF(Err.Number=0,2,1)
				On Error Goto 0
				If arrType = 1 Then
				'一维数组
					tpl "type", "数组"
					If size1 = -1 Then
						tpl.TagStr "table", GetTable(4)
						tpl "value", "此数组是空的，还没有任何元素"
					Else
						tpl.TagStr "table", GetTable(0)
						tpl "cname", "下标"
						tpl "cvalue", "元素值"
						tpl "count", size1+1
						tpl.Tag("top") = Easp.IIF(size1>49 And t=0,50,0)
						For i = 0 To size1
							If t = 0 And j>=50 Then Exit For
							tpl "name", i
							ShowValue o(i), "", 0
							tpl.Update "loop"
							j = j + 1
						Next
					End If
				Else
				'二维数组
					tpl "type", "二维数组"
					tpl "count", size2+1
					If t = 0 Then
						tpl.TagStr "table", GetTable(2)
						tpl "cno", "一维\二维"
						tpl.Tag("top") = Easp.IIF(size2>19, 20, 0)
						For j = 0 To size1
							tpl "field", "&nbsp;&nbsp;&nbsp;" & j & "&nbsp;&nbsp;&nbsp;"
							tpl.Update "col"
						Next
						For i = 0 To size2
							If i>20 Then Exit For
							tpl "i", i
							For j = 0 To size1
								'tpl "value", Easp.Str.HtmlEncode(o(j,i))
								ShowValue o(j,i), "", 0
								tpl.Update "fields"
							Next
							tpl.Update "rs"
						Next
					ElseIf t = 1 Then
						tpl.TagStr "table", GetTable(3)
						tpl "cname", "下标"
						tpl "cvalue", "元素值"
						For i = 0 To size2
							tpl "i", i+1
							For j = 0 To size1
								tpl "name", "("&j&", "&i&")"
								'tpl "value", Easp.Str.HtmlEncode(o(j,i))
								ShowValue o(j,i), "", 0
								tpl.Update "loop"
							Next
							tpl.Update "rs"
						Next
					End If
				End If
		End Select
	End Sub
	Private Function ShowValue(ByVal o, ByVal t, ByVal i)
		If IsObject(o) Then
			tpl.Tag("value") = "<span class=""info"">[ "&TypeName(o)&" Object ]" & Easp.IfThen(Easp.Has(t),", 要查看其中的内容，请使用 Easp.Ext(""Trace"")("&t&"("""&i&"""))") & "</span>"
		ElseIf IsArray(o) Then
			tpl.Tag("value") = "<span class=""info"">[ Array ]" & Easp.IfThen(Easp.Has(t),", 要查看其中的内容，请使用 Easp.Ext(""Trace"")("&t&"("""&i&"""))") & "</span>"
		Else
			tpl "value", Easp.Str.HtmlEncode(o)
		End If
	End Function
	Private Sub UpdateLoop(ByVal s)
		tpl "name", "<span class=""info"">"&s&"</span>"
		tpl "value", "<span class=""info"">"&Eval(s)&"</span>"
		tpl.Update "loop"
	End Sub
	Private Sub TraceDb(ByVal con, ByVal isall)
		If TypeName(con)<>"Connection" Then Exit Sub
		Dim t,i,j,k,db,f,s,arr1,arr2,dbtype : j = 0 : k = 0
		Set t = con.OpenSchema(20,Array(Empty,Empty,Empty,"TABLE"))
		arr1 = t.GetRows(-1)
		Set t = con.OpenSchema(20,Array(Empty,Empty,Empty,"VIEW"))
		arr2 = t.GetRows(-1)
		Easp.Db.Close(t)
		tpl "type", "数据库"
		If isall = 0 Then
			tpl.TagStr "table", GetTable(5) & GetTable(1)
			tpl "cno", "类型"
			tpl "cname", "名称"
			tpl "cvalue", "字段名"
			tpl "info", "您可以用 Easp.Ext(""Trace"")("":db.表名或视图名"") 查看单表或单视图的详细信息；用 Easp.Ext(""Trace"").TraceAll("":db"") 查看全部表和视图的详细信息。"
		ElseIf isall = 1 Then
			tpl.TagStr "table", GetTable(6)
		End If
		'Set db = Easp.db.New
		'db.Conn = con
		dbtype = Easp.Db.GetType(con)
		For i = 0 To Ubound(arr1,2)
			If isall = 0 Then
				tpl.Tag("no") = "表"
				tpl "name", arr1(2,i)
				'Set f = db.GR(arr1(2,i)&":1","1=-1","")
				Easp.Var("name") = arr1(2,i)
				Set f = Easp.Db.Execute(con, "Select Top 1 * From {=name} Where 1=-1")
				s = ""
				For j = 0 To f.Fields.Count-1
					s = s & ", " & f.Fields(j).Name
				Next
				tpl "value", Mid(s,2)
				Easp.Db.Close(f)
				tpl.Update "loop"
			ElseIf isall = 1 Then
				'此处如写成函数调用会出错，只能重复TraceTable函数
				Set t = con.OpenSchema(4,Array(Empty,Empty,arr1(2,i),Empty))
				tpl "tableorview", "数据表"
				tpl "name", Easp.IfThen(Easp.Has(t("TABLE_CATALOG")),"["&t("TABLE_CATALOG")&"].") & Easp.IfThen(Easp.Has(t("TABLE_SCHEMA")),"["&t("TABLE_SCHEMA")&"].") & "["&t("TABLE_NAME")&"]"
				While Not t.Eof
					tpl "field", t("COLUMN_NAME")
					If dbtype = "MSSQL" Then
						tpl "datatype", GetSQLDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"),t("CHARACTER_OCTET_LENGTH"),t("NUMERIC_PRECISION"),t("NUMERIC_SCALE"),t("DATETIME_PRECISION"))
					ElseIf dbtype = "ACCESS" Then
						tpl "datatype", GetACCDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"))
					End If
					tpl "nullable", Easp.IfThen(t("IS_NULLABLE"),"√")
					tpl "default", t("COLUMN_DEFAULT")
					tpl "desc", t("DESCRIPTION")
					tpl.Update "loop"
					t.MoveNext
				Wend
				tpl.Update "table"
				Easp.Db.Close(t)
			End If
			k = k + 1
		Next
		For i = 0 To Ubound(arr2,2)
			If isall = 0 Then
				tpl.Tag("no") = "视图"
				tpl "name", arr2(2,i)
				Easp.Var("name") = arr1(2,i)
				Set f = Easp.Db.Execute(con, "Select Top 1 * From {=name} Where 1=-1")
				'Set f = db.GR(arr2(2,i)&":1","1=-1","")
				s = ""
				For j = 0 To f.Fields.Count-1
					s = s & ", " & f.Fields(j).Name
				Next
				tpl "value", Mid(s,2)
				Easp.Db.Close(f)
				tpl.Update "loop"
			ElseIf isall = 1 Then
				'此处如写成函数调用会出错，只能再次重复TraceTable函数
				Set t = con.OpenSchema(4,Array(Empty,Empty,arr2(2,i),Empty))
				tpl "tableorview", "视图"
				tpl "name", Easp.IfThen(Easp.Has(t("TABLE_CATALOG")),"["&t("TABLE_CATALOG")&"].") & Easp.IfThen(Easp.Has(t("TABLE_SCHEMA")),"["&t("TABLE_SCHEMA")&"].") & "["&t("TABLE_NAME")&"]"
				While Not t.Eof
					tpl "field", t("COLUMN_NAME")
					If dbtype = "MSSQL" Then
						tpl "datatype", GetSQLDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"),t("CHARACTER_OCTET_LENGTH"),t("NUMERIC_PRECISION"),t("NUMERIC_SCALE"),t("DATETIME_PRECISION"))
					ElseIf dbtype = "ACCESS" Then
						tpl "datatype", GetACCDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"))
					End If
					tpl "nullable", Easp.IfThen(t("IS_NULLABLE"),"√")
					tpl "default", t("COLUMN_DEFAULT")
					tpl "desc", t("DESCRIPTION")
					tpl.Update "loop"
					t.MoveNext
				Wend
				tpl.Update "table"
				Easp.Db.Close(t)
			End If
			k = k + 1
		Next
		Easp.Db.Close(db)
		tpl "count", k
		Set t = Nothing
	End Sub
	'读取数据表信息(转载请保留版权) - Author:coldstone - 2010/10/19
	Private Sub TraceTable(ByVal tab, ByRef con)
		If TypeName(con)<>"Connection" Then Exit Sub
		Dim t,db,dbtype,dov
		Set t = con.OpenSchema(4,Array(Empty,Empty,tab,Empty))
		dov = con.OpenSchema(20,Array(Empty,Empty,tab,Empty))("TABLE_TYPE")
		tpl "tableorview", Easp.IIF(dov = "VIEW","视图","数据表")
		tpl "name", Easp.IfThen(Easp.Has(t("TABLE_CATALOG")),t("TABLE_CATALOG")&".") & Easp.IfThen(Easp.Has(t("TABLE_SCHEMA")),t("TABLE_SCHEMA")&".") & t("TABLE_NAME")
		dbtype = Easp.Db.GetType(con)
		While Not t.Eof
			tpl "field", t("COLUMN_NAME")
			If dbtype = "MSSQL" Then
				tpl "datatype", GetSQLDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"),t("CHARACTER_OCTET_LENGTH"),t("NUMERIC_PRECISION"),t("NUMERIC_SCALE"),t("DATETIME_PRECISION"))
			ElseIf dbtype = "ACCESS" Then
				tpl "datatype", GetACCDataType(t("DATA_TYPE"),t("COLUMN_FLAGS"),t("CHARACTER_MAXIMUM_LENGTH"))
			End If
			tpl "nullable", Easp.IfThen(t("IS_NULLABLE"),"√")
			tpl "default", t("COLUMN_DEFAULT")
			tpl "desc", t("DESCRIPTION")
			tpl.Update "loop"
			t.MoveNext
		Wend
		tpl.Update "table"
		Easp.Db.Close(t)
	End Sub
	'判断MSSQL数据类型及大小(coldstone呕心原创，转载请保留版权,@2010/10/19)
	Private Function GetSQLDataType(ByVal typeid, ByVal flag, ByVal maxlen, ByVal octlen, ByVal numpre, ByVal numscl, ByVal datepre)
		Dim tmp
		Select Case typeid
			Case 2 tmp = "smallint" & Easp.IfThen(flag=16,",自增")
			Case 3 tmp = "int" & Easp.IfThen(flag=20,",自增")
			Case 4 tmp = "real"
			Case 5 tmp = "float"
			Case 6 tmp = Easp.IIF(numpre=10,"smallmoney","money")
			Case 11 tmp = "bit"
			Case 12 tmp = "sql_variant"
			Case 17 tmp = "tinyint" & Easp.IfThen(flag=16,",自增")
			Case 20 tmp = "bigint" & Easp.IfThen(flag=16,",自增")
			Case 72 tmp = "uniqueidentifier"
			Case 128
				Select Case flag
					Case 116,20 tmp = "binaray(" & maxlen & ")"
					Case 230,134 tmp = "image"
					Case 624,528 tmp = "timestamp"
					Case 100,4 tmp = "varbinary(" & maxlen & ")"
					Case Else tmp = "未知binaray类型"
				End Select
			Case 129
				Select Case flag
					Case 116,20 tmp = "char(" & maxlen & ")"
					Case 230,134 tmp = "text"
					Case 100,4 tmp = "varchar(" & maxlen & ")"
					Case Else tmp = "未知char类型"
				End Select
			Case 130
				Select Case flag
					Case 116,20 tmp = "nchar(" & maxlen & ")"
					Case 230,134 tmp = "ntext"
					Case 100,4 tmp = "nvarchar(" & maxlen & ")"
					Case Else tmp = "未知nchar类型"
				End Select
			Case 131 tmp = "decimal/numeric(" & numpre & "," & numscl & ")"
			Case 135 tmp = Easp.IIF(datepre=0,"smalldatetime","datetime")
			Case Else tmp = "未知类型"
		End Select
		GetSQLDataType = tmp
	End Function
	'判断ACCESS数据类型及大小(coldstone呕心原创，转载请保留版权,@2010/10/19)
	Private Function GetACCDataType(ByVal typeid, ByVal flag, ByVal maxlen)
		Dim tmp
		Select Case typeid
			Case 2 tmp = "数字(整型)"
			Case 3 tmp = Easp.IIF(flag=90,"自动编号/数字(非空长整型)","数字(长整型)")
			Case 4 tmp = "数字(单精度型)"
			Case 5 tmp = "数字(双精度型)"
			Case 6 tmp = "货币"
			Case 7 tmp = "日期/时间"
			Case 11 tmp = "是/否"
			Case 17 tmp = "数字(字节)"
			Case 20 tmp = "bigint" & Easp.IfThen(flag=16,",自增")
			Case 72 tmp = "数字(同步复制ID)"
			Case 128 tmp = "OLE对象"
			Case 130
				Select Case flag
					Case 106,74 tmp = "文本(" & maxlen & ")"
					Case 234,202 tmp = "备注"
					Case Else tmp = "未知文本类型"
				End Select
			Case 131 tmp = "数字(小数)"
			Case Else tmp = "未知类型"
		End Select
		GetACCDataType = tmp
	End Function
End Class
%>