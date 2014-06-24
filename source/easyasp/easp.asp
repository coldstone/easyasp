<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><%
Option Explicit
'######################################################################
'## easp.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-04-23 11:19:19
'## Description :   EasyASP main class
'##
'######################################################################
Server.ScriptTimeOut = 300
Response.Buffer = True  '设置输出缓冲区
Dim Easp_Include_html
Dim Easp_Timer : Easp_Timer = Timer() '设置计时器
Dim Easp : Set Easp = New EasyASP '实例化Easp
%>
<!--#include file="easp.config.asp" -->
<%
Class EasyASP
  '定义Easp公共类
  Public Lang, [Error], Str, Var, Console, [Date], Db, Encrypt, Json, List, Fso, Http, Tpl, Upload, Cache, Xml
  '定义Easp预留公共接口
  Public Mo, A, B, C, D, E, F, G, H, I, J, K, M, L, N
  '定义私有变量
  Private o_rwt, o_ext
  Private s_basePath, s_pluginPath, s_cores, s_defaultPageName, s_charset
  Private b_debug
  Private i_timer, i_newId
  '构造函数
  Private Sub Class_Initialize()
    s_basePath       = "/easyasp/"
    s_pluginPath     = s_basePath & "plugin/"
    b_debug          = False
    s_defaultPageName  = "index.asp"
    s_charset        = "UTF-8"
    Response.Charset = s_charset
    i_timer          = Timer()
    i_newId          = 0
    Set o_rwt        = Server.CreateObject("Scripting.Dictionary")
    Set o_ext        = Server.CreateObject("Scripting.Dictionary")
    Set Lang         = Server.CreateObject("Scripting.Dictionary")
    Lang.CompareMode = 1
    s_cores          = "[Error], Str, Var, Console, [Date], Db, Encrypt, Json, List, Fso, Http, Tpl, Upload, Cache, Xml"
    Core_Do "on", s_cores
  End Sub
  '析构函数
  Private Sub Class_Terminate()
    Set o_rwt = Nothing
    ClearExt()
    Set o_ext = Nothing
    Core_Do "off", s_cores
    Set Lang = Nothing
  End Sub
  
  '设置和读取Easp路径配置
  Public Property Let BasePath(ByVal path)
    Dim pathNew
    pathNew = FixSitePath(path)
    If Str.IsSame(s_pluginPath, s_basePath & "plugin/") Then
      s_pluginPath = pathNew & "plugin/"
    End If
    s_basePath = pathNew
  End Property
  Public Property Get BasePath()
    BasePath = s_basePath
  End Property
  '设置和读取Easp插件文件夹路径配置
  Public Property Let PluginPath(ByVal path)
    s_pluginPath = FixSitePath(path)
  End Property
  Public Property Get PluginPath()
    PluginPath = s_pluginPath
  End Property
  '修正目录根路径
  Private Function FixSitePath(ByVal path)
    If Left(path,1) <> "/" Then path = "/" & path
    If Right(path,1) <> "/" Then path = path & "/"
    FixSitePath = path
  End Function
  
  '设置和读取是否开启调试模式
  Public Property Let [Debug](ByVal bool)
    b_debug = bool
  End Property
  Public Property Get [Debug]
    [Debug] = b_debug
  End Property
  
  '设置和读取默认首页文件名
  Public Property Let DefaultPageName(ByVal string)
    s_defaultPageName = string
  End Property
  Public Property Get DefaultPageName
    DefaultPageName = s_defaultPageName
  End Property
  
  Public Property Let [CharSet](ByVal string) '设置和读取文档编码
    s_charset = Ucase(string)
    Response.Charset = s_charset
  End Property
  Public Property Get [CharSet]()
    [CharSet] = s_charset
  End Property

  '创建或销毁核心类公共对象
  Private Sub Core_Do(ByVal t, ByVal s)
    Dim a_core, i : a_core = Split(s,", ")
    Select Case t
      Case "on"
        For i = 0 To Ubound(a_core)
          Execute "Set " & a_core(i) & " = New EasyASP_object"
        Next
      Case "off"
        For i = Ubound(a_core) To 0 Step -1
          Execute "Set " & a_core(i) & " = Nothing"
        Next
    End Select
  End Sub

  '输出字符串
  Public Sub Echo(ByVal s) : Response.Write s : End Sub
  '输出字符串和一个换行符
  Public Sub Print(ByVal s) : Response.Write s & VbCrLf : Response.Flush() : End Sub
  '输出字符串和一个html换行符
  Public Sub Println(ByVal s) : Response.Write s & "<br />" & VbCrLf : Response.Flush() : End Sub
  '输出字符串并将HTML标签转为普通字符
  Public Sub PrintHtml(ByVal s) : Response.Write Str.HtmlEncode(s) & VbCrLf : End Sub
  Public Sub PrintlnHtml(ByVal s) : Response.Write Str.HtmlEncode(s) & "<br />" & VbCrLf : End Sub
  '将任意变量直接输出为字符串(Json格式)
  Public Sub PrintString(ByVal s) : Response.Write Str.ToString(s) & VbCrLf : End Sub
  Public Sub PrintlnString(ByVal s) : Response.Write Str.ToString(s) & "<br />" & VbCrLf : End Sub
  '输出经过格式化的字符串
  Public Sub PrintFormat(ByVal s, ByVal f) : Response.Write Str.Format(s, f) & VbCrLf : End Sub
  Public Sub PrintlnFormat(ByVal s, ByVal f) : Response.Write Str.Format(s, f) & "<br />" & VbCrLf : End Sub
  '输出字符串并终止程序运行
  Public Sub PrintEnd(ByVal s) : Response.Write s : [Exit]() : End Sub
  '终止程序运行
  Public Sub [Exit]():  Set Easp = Nothing : Response.End() : End Sub

  '判断是否为空值
  Public Function isN(ByVal s)
    If IsEmpty(s) Or IsNull(s) Then IsN = True : Exit Function
    isN = False
    Select Case VarType(s)
      Case vbString
        If s = "" Then isN = True : Exit Function
      Case vbObject
        Select Case TypeName(s)
          Case "Nothing"
            isN = True : Exit Function
          Case "Recordset"
            If s.State = 0 Then isN = True : Exit Function
            If s.Bof And s.Eof Then isN = True : Exit Function
          Case "Dictionary"
            If s.Count = 0 Then isN = True : Exit Function
          Case "EasyASP_List"
            If s.Count = 0 Then isN = True : Exit Function
        End Select
      Case vbArray,8194,8204,8209
        If Ubound(s)=-1 Then isN = True : Exit Function
    End Select
  End Function
  '判断是否不为空值
  Public Function Has(ByVal s)
    Has = Not isN(s)
  End Function
  '判断三元表达式
  Public Function IIF(ByVal Cn, ByVal T, ByVal F)
    If Cn Then
      IIF = T
    Else
      IIF = F
    End If
  End Function
  '如果条件成立则返回某值，否则返回空值
  Public Function IfThen(ByVal Cn, ByVal T)
    IfThen = IIF(Cn,T,"")
  End Function
  '如果第1项不为空则返回第1项，否则返回第2项
  Public Function IfHas(ByVal v1, ByVal v2)
    IfHas = IIF(Has(v1), v1, v2)
  End Function

  '获取GET参数值
  '参数 queryString - "URL参数名[:为空时默认值]"
  Public Function [Get](ByVal queryString)
    Dim a_rwt, s_get, o_matches, s_default
    s_default = Str.GetColonValue(queryString)
    queryString = Str.GetColonName(queryString)
    '检测是否是伪静态
    a_rwt = IsRewriteRule()
    If a_rwt(0) Then
      '如果是伪静态则取出符合的参数值
      Set o_matches = Str.Match(a_rwt(2), queryString & "=(\$\d)")
      If o_matches.Count > 0 Then
        s_get = Str.Replace(a_rwt(3), a_rwt(1), o_matches(0).SubMatches(0))
      End If
      Set o_matches = Nothing
      If IsN(s_get) And Has(s_default) Then s_get = s_default
    Else
      '如果不是伪静态则取普通URL参数
      s_get = Request.QueryString(queryString)
      If Has(s_default) Then
        Dim i
        If Request.QueryString(queryString).Count > 1 Then
          s_get = ""
          For i = 1 To Request.QueryString(queryString).Count
            If i > 1 Then s_get = s_get & ", "
            s_get = s_get & IfHas(Request.QueryString(queryString)(i), s_default)
          Next
        Else
          s_get = IfHas(s_get, s_default)
        End If
      End If
    End If
    [Get] = s_get
  End Function
  '获取POST参数值
  '参数 formString - "表单项名[:为空时默认值]"
  Public Function Post(ByVal formString)
    Dim s_post, s_default
    '取出默认值
    s_default = Str.GetColonValue(formString)
    formString = Str.GetColonName(formString)
    If Upload.checkEntryType Then
    '如果是上传表单
      If Not Upload.IsUploaded Then Upload.GetData()
      Dim a_post, i
      If Upload.FormArray.Exists(formString) Then
      '如果表单项确实存在
        '取出表单项的值
        Set a_post = Upload.FormArray(formString)
        '如果是多项同名的表单，则分别取值并为空值赋默认值
        For i = 0 To a_post.Length - 1
          If i > 0 Then s_post = s_post & ", "
          s_post = s_post & IfHas(a_post(i), s_default)
        Next
      Else
      '如果表单项不存在直接退出
        Exit Function
      End If
    Else
    '如果是普通表单
      s_post = Request.Form(formString)
      If Has(s_default) Then
        If Request.Form(formString).Count > 1 Then
          s_post = ""
          For i = 1 To Request.Form(formString).Count
            If i > 1 Then s_post = s_post & ", "
            s_post = s_post & IfHas(Request.Form(formString)(i), s_default)
          Next
        Else
          s_post = IfHas(s_post, s_default)
        End If
      End If
    End If
    Post = s_post
  End Function

  '取页面地址
  Public Function GetUrl(ByVal param)
    Dim script_name, s_url, s_dir, s_rq, i_port
    Dim s_item, s_tmp, i, b_tmp, s_protocol, s_port
    script_name = Request.ServerVariables("SCRIPT_NAME")
    i_port = Request.ServerVariables("SERVER_PORT")
    s_rq = Request.QueryString()
    '取出当前页地址，如果是默认首页如index.asp则省略首页名
    s_url = Mid(script_name, 1, IIF(Str.IsSame(Right(script_name, Len("/" & s_defaultPageName)),"/" & s_defaultPageName), Len(script_name)-Len(s_defaultPageName), Len(script_name)))
    '取出所在站点目录路径
    s_dir  = Left(script_name,InstrRev(script_name,"/"))
    Select Case param
      Case "-3"  GetUrl = script_name '返回页面文件路径(-3)
      Case "-2"  GetUrl = s_dir '返回页面所在站点目录路径(-2)
      Case "-1", ""
        '取出包含主机名的页面地址
        If Request.ServerVariables("HTTPS")="on" Then
          s_protocol = "https://"
          s_port = IfThen(i_port <> 443, ":" & i_port)
        Else
          s_protocol = "http://"
          s_port = IfThen(i_port <> 80, ":" & i_port)
        End If
        s_url = s_protocol & Request.ServerVariables("SERVER_NAME") & s_port
        '返回主机名(-1)或者包含主机名的完整URL("")
        GetUrl = s_url & IfThen(isN(param), script_name & IfThen(Has(s_rq), "?" & s_rq))
      Case "0"  GetUrl = s_url & IfThen(Has(s_rq), "?" & s_rq) '返回页面站点URL带参数(0)
      Case "1"  GetUrl = script_name & IfThen(Has(s_rq), "?" & s_rq) '返回页面文件路径和URL参数(1)
      Case "-:all"  GetUrl = s_url '返回删除所有URL参数后的地址
      Case Else
        'URL参数处理
        If Has(s_rq) Then
            s_tmp = "" : i = 0 : param = "," & param & ","
            b_tmp = IIF(Str.IsIn(param,"-"),"Not InStr(param,"",-""&s_item&"","")>0","InStr(param,"",""&s_item&"","")>0")
            '处理URL参数白名单或黑名单
            For Each s_item In Request.QueryString()
              If Eval(b_tmp) Then
                If i > 0 Then s_tmp = s_tmp & "&"
                s_tmp = s_tmp & s_item & "=" & Request.QueryString(s_item)
                i = i + 1
              End If
            Next
            If Has(s_tmp) Then s_url = s_url & "?" & s_tmp
        End If
        GetUrl = s_url
    End Select
  End Function

  '取页面地址并带上新参数
  Public Function GetUrlWith(ByVal param, ByVal value)
    Dim s_url, s_page
    '取出新页面地址
    If Str.IsIn(param,"?") Then
      s_page = Str.GetName(param, "?")
      param  = Str.GetValue(param, "?")
    End If
    '取出页面地址和参数
    s_url = GetUrl(param)
    '带上新参数
    s_url = s_url & IfThen(Has(value),IIF(Str.IsIn(s_url, "?"), "&", "?") & value)
    '如果有新页面则替换为新页面地址
    If Has(s_page) Then
      s_url = s_page & IfThen(Str.IsIn(s_url, "?"), "?" & Str.GetValue(s_url, "?"))
    End If
    GetUrlWith = s_url
  End Function

  '替换Url参数
  Public Function ReplaceUrl(ByVal param, ByVal value)
    Dim a_rwt, o_matches
    a_rwt = IsRewriteRule()
    If a_rwt(0) Then
      '如果是伪静态页面
      Set o_matches = Str.Match(a_rwt(2), param & "=(\$\d)")
      If o_matches.Count > 0 Then
        ReplaceUrl = Str.ReplacePart(a_rwt(3), a_rwt(1), o_matches(0).SubMatches(0), value)
      Else
        ReplaceUrl = a_rwt(3)
      End If
      Set o_matches = Nothing
    Else
      ReplaceUrl = GetUrlWith("-" & param, param & "=" & value)
    End If
  End Function

  '伪静态规则设置（传统方法）
  Public Sub RewriteRule(ByVal rule, ByVal url)
    If (Left(rule,2)<>"^/" And Left(rule,1)<>"/") Or Instr(url,"?") = 0 Then Exit Sub
    o_rwt(NewID()) = Array(rule, url)
    '设置伪静态页面的Easp.Var("get.***")值
    Dim a_rwt, a_params, i, key
    a_rwt = IsRewriteRule()
    If a_rwt(0) Then
      a_params = Split(a_rwt(2), "&")
      For i = 0 To UBound(a_params)
        key = Str.GetName(a_params(i), "=")
        Var("get." & key) = [Get](key)
      Next
    End If
  End Sub
  '伪静态规则设置（推荐方法）
  Public Sub Rewrite(ByVal urlFile, ByVal rule, Byval urlParam)
    Dim a_filePath, i, a_tmp, s_rule, s_url
    '先去除规则中的^和$
    If Left(rule,1) = "^" Then rule = Mid(rule,2)
    If Right(rule,1) = "$" Then rule = Left(rule,Len(rule)-1)
    '如果页面地址为空，则默认为当前页，有两种状态（默认首页可能省略index.asp）
    urlFile = IfHas(urlFile, GetUrl("-:all") & "|" & GetUrl(-3))
    'urlFile参数可以包含多个页面地址，以|符号分隔
    a_tmp = Split(urlFile,"|")
    '处理每一个页面地址为一个单独的规则
    For i = 0 To Ubound(a_tmp)
      '组合rewrite规则
      s_rule = "^" & Str.RegexpEncode(a_tmp(i)) & "\?" & rule & "$"
      s_url = a_tmp(i) & "?" & urlParam
      RewriteRule s_rule, s_url
      'Console s_rule & " : " & s_url
    Next
  End Sub
  '检测本页面是否符合已设置的伪静态规则
  Public Function IsRewrite()
    IsRewrite = IsRewriteRule()(0)
  End Function
  Private Function IsRewriteRule()
    Dim b_rwt, s_rule, a_url, s_url, i
    Dim s_rwtRule, s_rwtGroup, s_rwtParam
    b_rwt = False
    s_rwtRule = ""
    s_rwtGroup = ""
    s_rwtParam = ""
    If Has(o_rwt) Then
      '和已经存储的伪静态规则进行比对
      s_url = GetUrl(0)
      For Each i In o_rwt
        s_rule = o_rwt(i)(0)
        If Str.Test(s_url, s_rule) Then
          '如果找到匹配则将相关规则存入一个数组
          b_rwt = True
          s_rwtParam = s_url
          s_rwtRule  = s_rule
          s_rwtGroup = Str.GetValue(o_rwt(i)(1),"?")
          Exit For
        End If
      Next
    End If
    '返回包含匹配规则的数组(是否Rewrite, 规则正则, 规则参数, 页面参数)
    IsRewriteRule = Array(b_rwt, s_rwtRule, s_rwtGroup, s_rwtParam)
  End Function
  
  '不缓存页面信息
  Public Sub NoCache()
    Response.Expires = 0
    Response.ExpiresAbsolute = Now() - 1
    Response.CacheControl = "no-cache"
    Response.AddHeader "Expires",Now() - 1
    Response.AddHeader "Pragma","no-cache"
    Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
  End Sub

  '为Dictionary设置键值
  Public Sub SetDictionaryKey(ByRef dict, ByVal key, ByVal value)
    If dict.Exists(key) Then
      dict(key) = value
    Else
      dict.Add key, value
    End If
  End Sub

  '服务器端跳转
  Public Sub RR(ByVal url)
    Response.Redirect(url)
  End Sub

  '获取用户IP地址
  Public Function GetIP()
    Dim addr, x, y
    x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    y = Request.ServerVariables("REMOTE_ADDR")
    addr = IIF(isN(x) or lCase(x)="unknown",y,x)
    If InStr(addr,".")=0 Then addr = "0.0.0.0"
    GetIP = addr
  End Function

  '服务器端生成唯一不重复编号
  '返回：String - 10位字符串
  Public Function NewID()
    Dim id
    If i_newId = 0 Then
      '生成一个时间戳
      i_newId = (DateDiff("s","1949-10-01",[Date].Format(Now(),"y-mm-dd"))+Timer())*100000
    End If
    i_newId = i_newId + 1
    id = i_newId
    NewID = NumberToString(id)
  End Function
  '十进制转为36进制
  Private Function NumberToString(n)
    Dim t(11), v, c, m, l
    c = 10
    v = n / 36
    Do While v > 0 And c > 0
      c = c - 1
      m = Int(n - Int(n / 36) * 36)
      t(c) = Chr(IIF(m<10, m+48, m+55))
      n = v
      v = n / 36
    Loop
    '加两位随机码
    'Randomize
    'l = Int(36 * Rnd)
    't(10) = Chr(IIF(l<10, l+48, l+55))
    'l = Int(36 * Rnd)
    't(11) = Chr(IIF(l<10, l+48, l+55))
    NumberToString = Join(t,"")
  End Function
  '批量生成不重复编号
  '返回：Array
  Public Function NewIDs(ByVal number)
    Dim a_tmp(), i
    ReDim a_tmp(number-1)
    For i = 0 To number-1
      a_tmp(i) = NewID
    Next
    NewIDs = a_tmp
  End Function
  
  '获取脚本执行时间（秒）
  Public Function GetScriptTime()
    GetScriptTime = GetScriptTimeByTimer(i_timer)
  End Function
  '获取以某个时间戳为开始的脚本执行时间（秒）
  Public Function GetScriptTimeByTimer(ByVal t)
    GetScriptTimeByTimer = FormatNumber((Timer()-t), 3, -1)
  End Function

  '设置一个Cookies值
  Public Sub SetCookie(ByVal name, ByVal value, ByVal config)
    Dim n,i,d_expires,s_domain,s_path,b_secure
    If isArray(config) Then
      For i = 0 To Ubound(config)
        If isDate(config(i)) Then
          d_expires = cDate(config(i))
        ElseIf Str.Test(config(i),"int") Then
          If config(i)<>0 Then d_expires = Now()+Int(config(i))/60/24
        ElseIf Str.Test(config(i),"domain") or Str.Test(config(i),"ip") Then
          s_domain = config(i)
        ElseIf Instr(config(i),"/")>0 Then
          s_path = config(i)
        ElseIf Str.IsInList("True,False,-1,0", config(i)) Then
          b_secure = CBool(config(i))
        End If
      Next
    Else
      If isDate(config) Then
        d_expires = cDate(config)
      ElseIf Str.Test(config,"int") Then
        If config<>0 Then d_expires = Now()+Int(config)/60/24
      ElseIf Str.Test(config,"domain") or Str.Test(config,"ip") Then
        s_domain = config
      ElseIf Instr(config,"/")>0 Then
        s_path = config
      ElseIf Str.IsInList("True,False,-1,0", config) Then
        b_secure = CBool(config)
      End If
    End If
    If Instr(name,">")>0 Then
      n = Str.GetValue(name,">")
      name = Str.GetName(name,">")
      Response.Cookies(name)(n) = value
    Else
      Response.Cookies(name) = value
    End If
    If Has(d_expires) Then Response.Cookies(name).Expires = d_expires
    If Has(s_domain) Then Response.Cookies(name).Domain = s_domain
    If Has(s_path) Then Response.Cookies(name).Path = s_path
    If Has(b_secure) Then Response.Cookies(name).Secure = b_secure
  End Sub

  '获取一个Cookies值
  Public Function Cookie(ByVal name)
    Dim p,t,coo
    If Instr(name,">") Then
      p = Str.GetName(name,">")
      name = Str.GetValue(name,">")
    End If
    If Has(p) And Has(name) Then
      If Request.Cookies(p).HasKeys Then
        coo = Request.Cookies(p)(name)
      End If
    ElseIf Has(name) Then
      coo = Request.Cookies(name)
    Else
      Cookie = "" : Exit Function
    End If
    If IsN(coo) Then Cookie = "": Exit Function
    Cookie = coo
  End Function

  '删除一个Cookies值
  Public Sub RemoveCookie(ByVal name)
    Dim s_name
    If Instr(name,">") > 0 Then
      s_name = Str.GetName(name,">")
      name = Str.GetValue(name,">")
    End If
    If Has(s_name) And Has(name) Then
      If Response.Cookies(s_name).HasKeys Then
        Response.Cookies(s_name)(name) = Empty
      End If
    ElseIf Has(name) Then
      Response.Cookies(name) = Empty
      Response.Cookies(name).Expires = Now() - 1
    End If
  End Sub

  '设置Application
  Public Sub SetApplication(ByVal key,ByRef value)
    Application.Lock
    If IsObject(value) Then
      Set Application(key) = value
    Else
      Application(key) = value
    End If
    Application.UnLock
  End Sub
  '获取Application
  Public Function GetApplication(ByVal key)
    If IsObject(Application(key)) Then
      Set GetApplication = Application(key)
    Else
      GetApplication = Application(key)
    End If
  End Function
  '删除Application
  Public Sub RemoveApplication(ByVal key)
    Application.Contents.Remove(key)
  End Sub
  '删除所有Application
  Public Sub RemoveAllApplication()
    Application.Contents.RemoveAll()
  End Sub
  
  '检测组件是否安装
  Public Function IsInstall(Byval s)
    On Error Resume Next : Err.Clear()
    IsInstall = False
    Dim obj : Set obj = Server.CreateObject(s)
    If Err.Number = 0 Then IsInstall = True
    Set obj = Nothing : Err.Clear()
  End Function


  '动态包含文件
  Public Sub Include(ByVal filePath)
    'On Error Resume Next
    ExecuteGlobal GetIncCode(IncRead(filePath), False)
    If Err.Number<>0 Then
      [error].Msg = " ( " & filePath & " )"
      [error].Raise 1
    End If
    Err.Clear()
  End Sub
  '得到动态包含文件运行的结果
  Public Function GetInclude(ByVal filePath)
    'On Error Resume Next
    ExecuteGlobal GetIncCode(IncRead(filePath), True)
    GetInclude = Easp_Include_html
    If Err.Number<>0 Then
      [error].Msg = " ( " & filePath & " )"
      [error].Raise 1
    End If
    Err.Clear()
  End Function

  '读取包含文件内容（无限级）
  Public Function IncRead(ByVal filePath)
    Dim s_content, s_rule, o_matchesInc, s_incFilePath, s_incContent,match
    s_content = Fso.Read(filePath)
    If isN(s_content) Then Exit Function
    s_content = Str.Replace(s_content, "<% *?@.*?%"&">","")
    s_content = Str.Replace(s_content, "(<%[^>]+?)(option +?explicit)([^>]*?%"&">)","$1'$2$3")
    s_rule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
    If Str.Test(s_content, s_rule) Then
      Set o_matchesInc = Str.match(s_content, s_rule)
      For Each match In o_matchesInc
        If LCase(match.SubMatches(0))="virtual" Then
          s_incFilePath = match.SubMatches(1)
        Else
          s_incFilePath = Mid(filePath, 1, InstrRev(filePath, IIF(Instr(filePath, ":")>0, "\", "/"))) & match.SubMatches(1)
        End If
        s_incContent = IncRead(s_incFilePath)
        s_content = Replace(s_content, match.Value, s_incContent)
      Next
      Set o_matchesInc = Nothing
    End If
    IncRead = s_content
  End Function
  '将文本内容转换为ASP代码
  Public Function GetIncCode(ByRef content, ByRef getHtml)
    'Original by Alan (alan[at]jobicn.com, author of EasyIDE)
    Dim s_tmp, s_code, s_codeTmp, s_codeBegin, i_startPos, i_endPos
    s_code = "" : i_startPos = 1 : i_endPos = Instr(content, "<%") + 2
    s_codeBegin = IIF(getHtml, "Easp_Include_html = Easp_Include_html & ", "Response.Write ")
    Do While i_endPos > i_startPos + 1
      s_tmp = Mid(content, i_startPos, i_endPos-i_startPos-2)
      i_startPos = Instr(i_endPos, content, "%"&">") + 2
      If Has(s_tmp) Then
        s_tmp = Replace(s_tmp, """", """""")
        s_tmp = Replace(s_tmp, vbCrLf, """ & vbCrLf & """)
        s_code = s_code & s_codeBegin & """" & s_tmp & """" & vbCrLf
      End If
      s_tmp = Mid(content, i_endPos, i_startPos-i_endPos-2)
      s_codeTmp = Str.Replace(s_tmp, "^\s*=\s*", s_codeBegin) & vbCrLf
      If getHtml Then
        s_codeTmp = Str.ReplaceLine(s_codeTmp, "^(\s*)response\.write([\( ])", "$1" & s_codeBegin & "$2") & vbCrLf
        s_codeTmp = Str.ReplaceLine(s_codeTmp, "^(\s*)Easp\.(Echo|Print|Println)([\( ])", "$1" & s_codeBegin & "$3") & vbCrLf
      End If
      s_code = s_code & s_codeTmp
      i_endPos = Instr(i_startPos, content, "<%") + 2
    Loop
    s_tmp = Mid(content, i_startPos)
    If Has(s_tmp) Then
      s_tmp = Replace(s_tmp,"""","""""")
      s_tmp = Replace(s_tmp,vbcrlf,""" & vbCrLf & """)
      s_code = s_code & s_codeBegin & """" & s_tmp & """" & vbCrLf
    End If
    If getHtml Then s_code = "Easp_Include_html = """" " & vbCrLf & s_code
    GetIncCode = Str.Replace(s_code, "(\n\s*\r)+", vbCrLf)
  End Function

  '加载插件
  Public Default Function Ext(ByVal name)
    Dim b_loaded, s_filePath
    name = Lcase(name) : b_loaded = True
    If Not o_ext.Exists(name) Then
      b_loaded = False
    Else
      If LCase(TypeName(o_ext(name))) <> "easyasp_" & name Then b_loaded = False
    End If
    If Not b_loaded Then
      s_filePath = s_pluginPath & "easp." & name & ".asp"
      If Fso.isFile(s_filePath) Then
        Include s_filePath
        Execute("Set o_ext(""" & name & """) = New EasyASP_" & name)
      Else
        If Easp.Debug Then
          '如果插件不存在则抛出异常
          [Error].FunctionName = "Easp.Ext(""" & name & """)"
          [Error].Detail = s_filePath
          [Error].Raise "error-easp-pluginpath"
        End If
      End If
    End If
    Set Ext = o_ext(name)
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

  '表单验证
  Public Function GetVal(ByVal string)
    Set GetVal = New EasyASP_Validation
    GetVal.Source = Easp.Get(string)
  End Function
  Public Function PostVal(ByVal string)
    Set PostVal = New EasyASP_Validation
    PostVal.Source = Easp.Post(string)
  End Function
  Public Function VarVal(ByVal string)
    Set VarVal = New EasyASP_Validation
    VarVal.Source = Easp.Var(string)
  End Function

  '将对象或者数组转换为Json字符串
  Public Function Encode(ByVal Object)
    Encode = Json.ToString(Object)
  End Function
  '将Json字符串解析为对象或者数组
  Public Function Decode(ByVal string)
    Set Decode = Json.Parse(string)
  End Function
  
  Public Sub Init() '初始化EasyASP
    Set [Error] = New EasyASP_Error
    [Error]("error-easp-pluginpath") = Lang("error-easp-pluginpath")
    Set Fso     = New EasyASP_Fso
    Set Str     = New EasyASP_String
    Set Console = New EasyASP_Console
    Set Var     = New EasyASP_Var
    Set [Date]  = New EasyASP_Date
    Set Db      = New EasyASP_Db
    Set Encrypt = New EasyASP_Encrypt
    Set Json    = New EasyASP_Json
    Set List    = New EasyASP_List
    Set Upload  = New EasyASP_MoLibUpload
    Set Http    = New EasyASP_Http
    Set Tpl     = New EasyASP_Tpl
    Set Cache   = New EasyASP_Cache
    Set Xml     = New EasyASP_Xml
  End Sub

End Class
Class EasyASP_object : End Class
%>
<!--#include file="core/easp.stringbuilder.asp"-->
<!--#include file="core/easp.error.asp"-->
<!--#include file="core/easp.validation.asp"-->
<!--#include file="core/easp.str.asp"-->
<!--#include file="core/easp.stringobject.asp"-->
<!--#include file="core/easp.var.asp"-->
<!--#include file="core/easp.console.asp"-->
<!--#include file="core/easp.date.asp"-->
<!--#include file="core/easp.db.asp"-->
<!--#include file="core/easp.encrypt.asp"-->
<!--#include file="core/easp.json.asp"-->
<!--#include file="core/easp.list.asp"-->
<!--#include file="core/easp.fso.asp"-->
<!--#include file="core/easp.upload.asp"-->
<!--#include file="core/easp.http.asp"-->
<!--#include file="core/easp.tpl.asp"-->
<!--#include file="core/easp.cache.asp"-->
<!--#include file="core/easp.xml.asp"-->
<%'页面加载完毕后销毁Easp实例%>
<script language="vbscript" runat="server">If TypeName(Easp) = "EasyASP" Then Set Easp = Nothing</script>