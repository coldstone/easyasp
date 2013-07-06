<%
'#################################################################################
'##	easp.upload.asp
'##	------------------------------------------------------------------------------
'##	Feature		:	EasyAsp Upload Class
'##	Version		:	v2.2 Alpha
'##	Author		:	Coldstone(coldstone[at]qq.com)
'##	Update Date	:	2010/08/13 12:15:14
'##	Description	:	Upload file(s) with EasyASP
'#################################################################################
Dim EasyAsp_o_updata
Class EasyAsp_Upload
	Public Form, File, Count
	Private s_charset,s_key,s_allowed,s_denied
	Private s_savePath,s_jsonPath,s_progressPath,s_progExt
	Private i_fileMaxSize,i_totalMaxSize,i_blockSize
	Private b_useProgress, b_autoMD,b_random, o_fso, o_db
	'构造函数
	Private Sub Class_Initialize
		'默认编码，继承自Easp
		s_charset	= Easp.CharSet
		'默认此次上传唯一ID，无（用于进度条）
		s_key		= ""
		'默认仅允许上传类型，所有文件
		s_allowed	= ""
		'默认不允许上传类型，没有
		s_denied	= ""
		'默认文件保存位置，当前目录
		s_savePath	= ""
		'默认自动建立不存在的文件夹
		b_autoMD	= True
		'默认不使用随机文件名
		b_random	= False
		'默认单个文件允许大小，不限制
		i_fileMaxSize	= 0
		'默认总文件大小，不限制
		i_totalMaxSize = 0
		'分块上传默认每次上传64K
		i_blockSize = 64 * 1024
		b_useProgress = False
		'默认进度条临时文件夹，根目录下的/__uptemp/目录
		s_progressPath = "/__uptemp/"
		s_jsonPath	= ""
		s_progExt = ".txt"
		Easp.Error(71) = "表单类型错误，表单只能是""multipart/form-data""类型！"
		Easp.Error(72) = "请先选择要上传的文件！"
		Easp.Error(73) = "上传文件失败，上传文件总大小超过了限制！"
		Easp.Error(74) = "上传文件失败，上传文件不能为空！"
		Easp.Error(75) = "上传文件失败，文件大小超过了限制！"
		Easp.Error(76) = "上传文件失败，不允许上传此类型的文件！"
		Easp.Error(77) = "上传文件失败！"
		Easp.Error(78) = "获取文件失败！"
		Easp.Error(79) = "本次上传KEY不能为空，否则上传进度条不可用！"
		Easp.Error(70) = "保存进度条目录必须以 / 开头！"
		Set Form = Server.CreateObject("Scripting.Dictionary")
		Set File = Server.CreateObject("Scripting.Dictionary")
		Count = 0
		Set o_db = Easp.db.New
		GetConn Easp.db.Conn
	End Sub
	
	'设置文件编码
	Public Property Let CharSet(ByVal s)
		s_charset = UCase(s)
	End Property
	'设置数据库连接对象
	Public Property Let Conn(ByVal o)
		GetConn o
	End Property
	Public Property Get Conn
		If TypeName(o_db.Conn) = "Connection" Then
			Set Conn = o_db.Conn
		Else
			If IsObject(Conn) Then Set Conn = Nothing
			Easp.Error.Raise 13
		End If
	End Property
	Private Sub GetConn(ByVal o)
		If TypeName(o) = "Connection" Then
			If o.State = 1 Then
				o_db.Conn = o
			Else
				Easp.Error.Raise 13
			End If
		End If
	End Sub
	'设置进度条Json文件保位置：
	Public Property Let Key(ByVal s)
		If Not b_useProgress Then Exit Property
		If Easp.IsN(s) Then Easp.Error.Raise 79 : Exit Property
		s_key = s
		s_jsonPath = absPath(s_progressPath) & s & s_progExt
	End Property
	'生成上传唯一KEY
	Public Property Get GenKey
		GenKey = "EASPUP-" & Easp.DateTime(Now,"ymmddhhiiss") & Easp.RandStr("-<16>:0123456789ABCDEF")
	End Property
	'设置单个文件最大尺寸
	Public Property Let FileMaxSize(ByVal n)
		i_fileMaxSize = n * 1024
	End Property
	Public Property Get FileMaxSize
		FileMaxSize = i_fileMaxSize / 1024
	End Property
	'设置文件最大总尺寸
	Public Property Let TotalMaxSize(ByVal n)
		i_totalMaxSize = n * 1024
	End Property
	Public Property Get TotalMaxSize
		TotalMaxSize = i_totalMaxSize / 1024
	End Property
	'设置允许上传的文件类型，用"|"分隔
	Public Property Let Allowed(ByVal s)
		s_allowed = Replace(s,",","|")
	End Property
	Public Property Get Allowed
		Allowed = s_allowed
	End Property
	'设置禁止上传的文件类型，用"|"分隔
	Public Property Let Denied(ByVal s)
		s_denied = Replace(s,",","|")
	End Property
	Public Property Get Denied
		Denied = s_denied
	End Property
	'设置文件上传后保存的路径
	Public Property Let SavePath(ByVal s)
		Dim Matches,Match,t
		If Easp.Test(s,"<.+?>") Then
			Set Matches = Easp.RegMatch(s,"<(.+?)>")
			For Each Match In Matches
				t = Easp.DateTime(Now,Match.SubMatches(0))
				s = Replace(s,Match.Value,t)
			Next
		End If
		If Not Instr(s,":") = 2 Then
			If Right(s,1) <> "/" Then s = s & "/"
		Else
			If Right(s,1) <> "\" Then s = s & "\"
		End If
		s_savepath = s
	End Property
	Public Property Get SavePath
		SavePath = absPath(s_savepath)
	End Property
	'设置是否使用进度条
	Public Property Let UseProgress(ByVal b)
		b_useProgress = b
		If b Then 
			Easp.Use "Fso"
			Set o_fso = New EasyAsp_Fso
		End If
	End Property
	'设置保存进度条临时文件的文件夹
	Public Property Let ProgressPath(ByVal s)
		If Easp.IsN(s) Then Exit Property
		If Left(s,1)<>"/" Then Easp.Error.Raise 70 : Exit Property
		If Right(s,1)<>"/" Then s = s & "/"
		s_progressPath = s
		If Easp.Has(s_key) Then s_jsonPath = absPath(s_progressPath) & s_key & s_progExt
	End Property
	Public Property Get ProgressPath
		ProgressPath = s_progressPath
	End Property
	'获取进度条临时文件的Web地址
	Public Function ProgressFile(ByVal key)
		If Easp.Has(key) Then
			ProgressFile = s_progressPath & key & s_progExt
		End If
	End Function
	'设置是否自动创建不存在的文件夹
	Public Property Let AutoMD(ByVal b)
		b_autoMD = b
	End Property
	'设置是否重命名上传文件为随机文件名
	Public Property Let Random(ByVal b)
		b_random = b
	End Property
	'设置分块上传大小，单位K
	Public Property Let BlockSize(ByVal i)
		i_blockSize = Int(i) * 1024
	End Property
	'私有方法：取目录绝对路径
	Private Function absPath(ByVal s)
		If Easp.IsN(s) Then s = "."
		If Instr(s,":") <> 2 Then s = Server.MapPath(s)
		If Right(s,1)<>"\" Then s = s & "\"
		absPath = s
	End Function
	'私有方法：检测文件是否是允许上传类型
	Public Function checkFileType(ByVal t)
		checkFileType = True
		'如果设置了仅允许上传类型
		If Easp.Has(s_allowed) Then
			If Not Easp.Test(t, "^" & s_allowed & "$") Then
				checkFileType = False
				Exit Function
			End If
		'如果没有设置仅允许上传类型并设置了不允许上传类型
		ElseIf Easp.Has(s_denied) Then
			If Easp.Test(t,"^" & s_denied & "$") Then
				checkFileType = False
				Exit Function
			End If
		End If
	End Function
	'开始上传本地数据：
	Public Sub StartUpload
		If TypeName(o_db.Conn) <> "Connection" Then
				GetConn Easp.db.Conn
		End If
		'检测表单是否multipart/form-data类型
		Dim FormType : FormType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")
		If LCase(FormType(0)) <> "multipart/form-data" Then
			Easp.Error.Raise 71
			Exit Sub
		End If
		Dim o_strm, o_prog, o_file
		Dim s_block, s_blockData, s_start, s_formName, s_formValue, s_fileName, s_data
		Dim i_total, i_loaded, i_block, i_formStart, i_formEnd, i_Start, i_End, i_dataStart, i_dataEnd
		Dim CrLf
		'取得表单总大小
		i_total = Request.TotalBytes
		'如果表单的内容为空，则退出上传程序
		If i_total < 1 Then Easp.Error.Raise 72 : Exit Sub
		Set o_strm = Server.CreateObject("ADODB.Stream")
		'临时数据储存区
		Set EasyAsp_o_updata = Server.CreateObject("ADODB.Stream")
		EasyAsp_o_updata.Type = 1
		EasyAsp_o_updata.Mode =3
		EasyAsp_o_updata.Open
		'已读取的大小
		i_loaded = 0
		'如果使用进度条
		If b_useProgress Then
			'创建进度条文件夹
			Easp.Use "Fso"
			If Not Easp.Fso.IsFolder(s_progressPath) Then
				Easp.Fso.MD s_progressPath
			End If
			'创建Json文件记录进度条数据
			Set o_prog = New Easp_Upload_Progress
			'写入上传总大小到Json文件
			o_prog.TotalSize = i_total
			'创建Json文件
			o_prog.Create(s_jsonPath)
		End If
		'循环分块读取
		Do While i_loaded < i_total
			i_block = i_blockSize
			If i_block + i_loaded > i_total Then i_block = i_total - i_loaded
			s_block = Request.BinaryRead(i_block)
			i_loaded = i_loaded + i_block
			'写入分块数据
			EasyAsp_o_updata.Write s_block
			'更新进度条数据
			If b_useProgress Then o_prog.Update(i_loaded)
		Loop
		'EasyAsp_o_updata.Write  Request.BinaryRead(i_total)
		'将数据块读出处理
		EasyAsp_o_updata.Position = 0
		s_blockData = EasyAsp_o_updata.Read
		i_formStart = 1
		i_formEnd = LenB(s_blockData)
		CrLf = chrB(13) & chrB(10)
		s_start = MidB(s_blockData,1, InStrB(i_formStart,s_blockData,CrLf)-1)
		i_start = LenB(s_start)
		i_formStart = i_formStart + i_start + 1
		While (i_formStart + 10) < i_formEnd 
			i_End = InStrB(i_formStart,s_blockData,CrLf & CrLf)+3
			o_strm.Type = 1
			o_strm.Mode =3
			o_strm.Open
			EasyAsp_o_updata.Position = i_formStart
			EasyAsp_o_updata.CopyTo o_strm, i_End-i_formStart
			o_strm.Position = 0
			o_strm.Type = 2
			o_strm.Charset = s_charset
			s_data = o_strm.ReadText
			o_strm.Close
			'Easp.w s_data
			'取得表单项目名称
			i_formStart = InStrB(i_End,s_blockData,s_start)
			i_dataStart = InStr(22,s_data,"name=""",1) + 6
			i_dataEnd = InStr(i_dataStart,s_data,"""",1)
			s_formName = Mid(s_data,i_dataStart,i_dataEnd-i_dataStart)
			'如果是文件
			If InStr(43,s_data,"filename=""",1) > 0 Then
				Set o_file = New Easp_Upload_FileInfo
				o_file.autoMD = b_autoMD
				'取得文件大小
				o_file.Size = i_formStart - i_End - 3
				'如果文件大小超过了限制
				If (i_fileMaxSize>0 And (o_file.Size)>i_fileMaxSize) Then
					o_file.isSize = False
					Easp.Error.Raise 75
				End If
				'如果不为空
				If o_file.Size > 0 Then
					'取得文件名
					i_dataStart = InStr(i_dataEnd,s_data,"filename=""",1) + 10
					i_dataEnd = InStr(i_dataStart,s_data,"""",1)
					s_fileName = Mid(s_data,i_dataStart,i_dataEnd-i_dataStart)
					o_file.Client = s_fileName
					o_file.OldPath = Left(s_fileName, InstrRev(s_fileName, "\"))
					o_file.NewPath = absPath(s_savepath)
					o_file.WebPath = s_savepath
					o_file.Name = Mid(s_fileName, InstrRev(s_fileName, "\")+1)
					o_file.Ext = Mid(o_file.Name, InstrRev(o_file.Name,".")+1)
					o_file.NewName = Easp.IIF(b_random,Easp.DateTime(Now,"ymmddhhiiss")&Easp.RandStr("<100000-999999>") & "." & o_file.Ext,o_file.Name)
					'如果文件类型不允许
					If Not checkFileType(o_file.Ext) Then
						o_file.isType = False
						Easp.Error.Raise 76
					End If
					'取得MIME类型
					i_dataStart = InStr(i_dataEnd,s_data,"Content-Type: ",1) + 14
					i_dataEnd = InStr(i_dataStart,s_data,vbCr)
					o_file.MIME = Mid(s_data,i_dataStart,i_dataEnd-i_dataStart)
					o_file.Start = i_End
					o_file.FormName = s_formName
					'文件大小和类型都正确则上传文件总数加1
					If o_file.isSize And o_file.isType Then
						Count = Count + 1
					End If
				End If
				If NOT File.Exists(s_formName) Then
					'添加到文件集合
					File.Add s_formName, o_file
				End If
				Set o_file = Nothing
			Else
				'如果是表单项目
				o_strm.Type = 1
				o_strm.Mode = 3
				o_strm.Open
				EasyAsp_o_updata.Position = i_End 
				EasyAsp_o_updata.CopyTo o_strm, i_formStart-i_End-3
				o_strm.Position = 0
				o_strm.Type = 2
				o_strm.Charset = s_charset
				s_formValue = o_strm.ReadText
				o_strm.Close
				'添加到表单集合
				If Form.Exists(s_formName) Then
					Form(s_formName) = Form(s_formName) & ", " & s_formValue
				Else
					Form.Add s_formName, s_formValue
				End If
			End If
			i_formStart = i_formStart + i_start + 1
		Wend
		s_blockData = ""
		Set o_strm = Nothing
		Set o_prog = Nothing
	End Sub
	'保存全部文件
	Public Sub SaveAll
		Dim f
		If Easp.Has(File) Then
			For Each f In File
				File(f).Save
			Next
		End If
	End Sub
  '析构函数
  Private Sub Class_Terminate  
		If Request.TotalBytes > 0 Then
			Form.RemoveAll
			File.RemoveAll
			Easp.C(EasyAsp_o_updata)
		End If
		Set Form = Nothing
		Set File = Nothing
		Set o_db = Nothing
		'删除临时进度条文件
		If b_useProgress Then
			If o_fso.IsFile(s_jsonPath) Then o_fso.DelFile s_jsonPath
			Set o_fso = Nothing
		End If
  End Sub
End Class
'上传文件信息
Class Easp_Upload_FileInfo
	Public FormName, Client, OldPath, NewPath, WebPath, Name, NewName, Ext, Size, MIME
	Public isSize, isType, autoMD, Start
	Private Sub Class_Initialize
		'表单项名称
		FormName = ""
		'客户端文件位置
		Client = ""
		'客户端文件路径
		OldPath = ""
		'服务器端保存物理路径
		NewPath = ""
		'服务器端保存Web路径
		WebPath = ""
		'原文件名称
		Name = ""
		'新文件名称
		NewName = ""
		'文件扩展名
		Ext = ""
		'文件大小
		Size = 0
		'文件在数据块中的开始位置
		Start = 0
		'文件MIME类型
		MIME = ""
		'判断文件大小和类型
		isSize = True
		isType = True
		autoMD = True
	End Sub
	'文件保存为
	Public Function SaveAs(ByVal p)
		Dim o_strm,s_path
		SaveAs = True
		'如果上传文件大小为空或超过抛出异常
		If Size <= 0 Then
			SaveAs = False
			Exit Function
		ElseIf Not isSize Then
			SaveAs = False
			Easp.Error.Raise 75
			Exit Function
		End If
		'如果上传的不是文件或保存路径为空
		If Easp.IsN(p) Or Easp.IsN(Name) Or Start = 0 Or Right(p,1)="/" Then
			SaveAs = False
			Exit Function
		End If
		'如果文件类型不允许
		If Not isType Then
			SaveAs = False
			Easp.Error.Raise 76
			Exit Function
		End If
		If autoMD Then
			Easp.Use "Fso"
			s_path = Left(p,InstrRev(p,"\"))
			If Not Easp.Fso.IsFolder(s_path) Then Easp.Fso.MD(s_path)
		End If
		'保存文件
		Set o_strm = Server.CreateObject("Adodb.Stream")
		o_strm.Mode = 3
		o_strm.Type = 1
		o_strm.Open
		EasyAsp_o_updata.position = Start
		EasyAsp_o_updata.copyto o_strm, Size
		o_strm.SaveToFile p, 2
		o_strm.Close
		Set o_strm = Nothing
		'Easp.Upload.DelProgressFile()
	End Function
	'保存(自动选择地址)
	Public Function Save
		Save = SaveAs(NewPath & NewName)
	End Function
End Class
'进度条生成类
Class Easp_Upload_Progress
	Private s_path,i_total
	Private o_json,o_timer
	Private Sub Class_Initialize
		Easp.Use "Fso"
		'覆盖已有文件必须打开
		Easp.Fso.OverWrite = True
		i_total = 0
		s_path = ""
	End Sub
	Private Sub Class_Terminate
		If TypeName(o_json)="EasyAsp_JSON" Then Set o_json = Nothing
	End Sub
	'上传文件的总大小
	Public Property Let TotalSize(ByVal i)
		i_total = i
	End Property
	'创建Json文件
	Public Sub Create(ByVal p)
		s_path = p
		o_timer = Timer
		Easp.Use "Json"
		'建立EaspJSon对象
		Set o_json = Easp.Json.New(0)
		o_json("total") 	= Easp.Fso.FormatSize(i_total,"AUTO")
		o_json("uploaded")	= "0 KB"
		o_json("percent")	= "0"
		o_json("speed") 	= "0 KB"
		o_json("passtime") 	= "00:00:00"
		o_json("totaltime")	= "00:00:00"
		o_json("uploadtime")= Easp.DateTime(Now(),"y-mm-dd hh:ii:ss")
		Call Easp.Fso.CreateFile(s_path, o_json.jsString)
	End Sub
	'更新进度条Json文件
	Sub Update(ByVal loaded)
		Dim speed,cTimer,totalTime,remainTime,percent
		speed = 0.0001
		cTimer = Timer
		'计算上传速度
		If (cTimer - o_timer)>0 Then speed = loaded / (cTimer - o_timer)
		'计算总时间
		totalTime = i_total / speed
		'计算剩余时间
		remainTime = (i_total - loaded) / speed
		'计算上传百分比
		percent = Round(loaded *100 / i_total,1)
		'更新Json文件
		o_json("uploaded")	= Easp.Fso.FormatSize(loaded,"AUTO")
		o_json("percent")	= percent
		o_json("speed") 	= Easp.Fso.FormatSize(speed,"AUTO") & "/S" 
		o_json("totaltime") = SecToTime(totalTime)
		o_json("remaintime")= SecToTime(remainTime)
		o_json("uploadtime")= Easp.DateTime(Now(),"y-mm-dd hh:ii:ss")
		Call Easp.Fso.CreateFile(s_path, o_json.jsString)       
	End Sub
	'将秒转换为标准格式时间
	private Function SecToTime(ByVal sec)
		'On Error Resume Next
		Dim h : h = "00"
		Dim m : m = "00"
		Dim s : s = "00"
		If isNumeric(sec) Then
			h = Right("0" & Round(sec/3600), 2)
			m = Right("0" & Round(mod_(sec,3600) / 60), 2)
			s = Right("0" & Round(mod_(sec,60)), 2)
		End If
		SecToTime = (h & ":" & m & ":" & s)
	End Function
	Private Function mod_(ByVal a, ByVal b)
		mod_ = (a - int(a/b)*b)
	End Function
End Class
%>