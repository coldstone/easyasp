<%
'######################################################################
'## easp.http.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP XMLHTTP Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-05-24 08:12:24
'## Description :   Request XMLHttp Data in EasyASP
'## 
'######################################################################
Class EasyASP_Http
  Public Method, CharSet, Async, User, Password, Html, Headers, Body, Text, SaveRandom
  Public ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
  Private s_data, s_url, s_ohtml, o_rh
  
  Private Sub Class_Initialize
    Easp.Error("error-http-object") = Easp.Lang("error-http-object")
    Easp.Error("error-http-serverdown") = Easp.Lang("error-http-serverdown")
    Easp.Error("error-http-status") = Easp.Lang("error-http-status")
    Easp.Error("error-http-remote") = Easp.Lang("error-http-remote")
    Easp.Error("error-http-wrongstart") = Easp.Lang("error-http-wrongstart")
    Easp.Error("error-http-wrongend") = Easp.Lang("error-http-wrongend")
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
    'Easp.Use "List"
    Set o_rh = Easp.Json.NewObject
'    ReDim a_rh(-1)
  End Sub
  
  Private Sub Class_Terminate
    Set o_rh = Nothing
  End Sub

  '建新Easp远程文件操作类实例
  Public Function [New]()
    Set [New] = New EasyASP_Http
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
        n = Trim(Easp.Str.GetColonName(a(i)))
        v = Trim(Easp.Str.GetColonValue(a(i)))
        o_rh(n) = v
      Next
    Else
      n = Trim(Easp.Str.GetColonName(a))
      v = Trim(Easp.Str.GetColonValue(a))
      o_rh(n) = v
    End If
  End Sub
  '设置或获取单项请求头信息
  Public Property Let RequestHeader(ByVal name, ByVal value)
    o_rh(name) = value
  End Property
  Public Property Get RequestHeader(ByVal name)
    If Easp.Has(name) Then
      RequestHeader = o_rh(name)
    Else
      Dim dic, key, s
      Set dic = o_rh.GetDictionary
      For Each key In dic
        s = s & key & ":" & dic(key) & vbCrLf
      Next
      Set dic = Nothing
      RequestHeader = s
    End If
  End Property

  '配置URL
  Public Property Let Url(ByVal string)
    s_url = string
  End Property
  
'  '设置http的RequestHeader
  Private Sub SetHeaderTo(ByRef ht)
    Dim dic, key
    Set dic = o_rh.GetDictionary
    'Easp.Console dic
    For Each key In dic
      ht.setRequestHeader key, dic(key)
      'Easp.Console key & "/" & dic(key)
    Next
    Set dic = Nothing
  End Sub
  
  '属性配置模式下打开连接远程
  Public Function [Open]
    [Open] = GetData(s_url, Method, Async, s_data, User, Password)
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
    Dim a_http, i, ht, chru, s_serData
    a_http = Split("Msxml2.ServerXMLHTTP.6.0 MSXML2.serverXMLHTTP MSXML2.XMLHTTP Microsoft.XMLHTTP")
    i = 0
    '抓取地址
    If Easp.IsN(uri) Then Exit Function
    '通过URL临时指定编码
    If Easp.Str.Test(uri,"^[\w\d-]+>https?://") Then
      CharSet = Easp.Str.GetName(uri,">")
      uri = Easp.Str.GetValue(uri,">")
    End If
    s_url = uri
    '方法：POST或GET
    m = Easp.IIF(Easp.Has(m),UCase(m),"GET")
    '异步
    If Easp.IsN(async) Then async = False
    '构造Get传数据的URL
    If Easp.Has(data) Then s_serData = Serialize__(data)
    If m = "GET" And Easp.Has(data) Then uri = uri & Easp.IIF(Instr(uri,"?")>0, "&", "?") & s_serData
    'Easp.Console m & "/" & uri & "/" & async & "/" & u & "/" & p
    Do While True
      On Error Resume Next
      If Easp.isInstall(a_http(i)) Then
        Set ht = Server.CreateObject(a_http(i))
        If Instr(a_http(i), "server") Then
        '设置超时时间
          ht.SetTimeOuts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
        End If
        '打开远程页
        If Easp.Has(u) Then
          '如果有用户名和密码
          ht.open m, uri, async, u, p
        Else
          '匿名
          ht.open m, uri, async
        End If
        If m = "POST" Then
          If Not o_rh.Has("Content-Type") Then
            o_rh("Content-Type") = "application/x-www-form-urlencoded"
          End If
          SetHeaderTo ht
          '有发送的数据
          ht.send s_serData
        Else
          SetHeaderTo ht
          ht.send
        End If
        'Easp.Console a_http(i)
        If Err.Number = 0 Then
          'Easp.Console a_http(i)
          Exit Do
        End If
      End If
      If i < Ubound(a_http) Then
        i = i + 1
      Else
        If Easp.Debug Then
          Easp.Error.FunctionName = "Http.GetData"
          Easp.Error.Raise "error-http-object"
        End If
        Exit Do
      End If
      On Error Goto 0
    Loop
    '检测返回数据
    If ht.readyState <> 4 Then
      GetData = "error:server is down"
      Set ht = Nothing
      If Easp.Debug Then
        Easp.Error.FunctionName = "Http.GetData"
        Easp.Error.Detail = uri
        Easp.Error.Raise "error-http-serverdown"
      End If
      Exit Function
    ElseIf ht.Status = 200 Then
      Headers = ht.getAllResponseHeaders()
      Body = ht.responseBody
      If Easp.IsN(CharSet) Then
        Text = ht.responseText
        '从Header中提取编码信息
        If Easp.Str.Test(Headers,"charset=([\w-]+)") Then
          CharSet = Easp.Str.Replace(Headers,"([\s\S]+)charset=([\w-]+)([\s\S]+)","$2")
        '如果是Xml文档，从文档中提取编码信息
        ElseIf Easp.Str.Test(Headers,"Content-Type: ?text/xml") Then
          CharSet = Easp.Str.Replace(Text,"^<\?xml\s+[^>]+encoding\s*=\s*""([^""]+)""[^>]*\?>([\s\S]+)","$1")
        '从文件源码中提取编码
        ElseIf Easp.Str.Test(Text,"<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>") Then
          CharSet = Easp.Str.Replace(Text,"([\s\S]+)<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>([\s\S]+)","$2")
        End If
        '如果无法获取远程页的编码则继承Easp的编码设置
        If Easp.IsN(CharSet) Then CharSet = "UTF-8"
      End If
      GetData = Bytes2Bstr__(Body, CharSet)
    Else
      GetData = "error:" & ht.Status & " " & ht.StatusText
      If Easp.Debug Then
        Easp.Error.FunctionName = "Http.GetData"
        Easp.Error.Detail = Array(uri, ht.Status)
        Easp.Error.Raise "error-http-status"
      End If
    End If
    If Err.Number > 0 Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "Http.GetData"
        Easp.Error.Detail = Array(uri, m)
        Easp.Error.Raise "error-http-remote"
      End If
    End If
    Set ht = Nothing
    s_ohtml = GetData
    Html = s_ohtml
  End Function
    
  '按正则查找返回HTML中符合的第一个字符串
  Public Function Find(ByVal rule)
    Find = FindString(s_ohtml, rule)
  End Function
  '按正则在字符串中查找符合的第一个子字符串
  Public Function FindString(ByVal s, ByVal rule)
    If Easp.Str.Test(s,rule) Then FindString = Easp.Str.Replace(s,"([\s\S]*)("&rule&")([\s\S]*)","$2")
  End Function
  
  '按正则查找返回HTML中符合的第一个字符串并选择编组
  '可按正则编组选择其中的一部分
  Public Function [Select](ByVal rule, ByVal part)
    [Select] = SelectString(s_ohtml, rule, part)
  End Function
  '按正则查找字符串中符合的第一个子字符串并选择编组
  Public Function SelectString(ByVal s, ByVal rule, ByVal part)
    If Easp.Str.Test(s,rule) Then
      '$0匹配字符串本身
      part = Replace(part,"$0",FindString(s,rule))
      '按正则编组分别替换
      SelectString = Easp.Str.Replace(s,"(?:[\s\S]*)(?:"&rule&")(?:[\s\S]*)",part)
    End If
  End Function
  
  '按正则查找返回HTML中符合的字符串组，返回数组
  Public Function Search(ByVal rule)
    Search = SearchString(s_ohtml, rule)
  End Function
  '按正则查找字符串中符合的子字符串组，返回数组
  Public Function SearchString(ByVal s, ByVal rule)
    Dim matches,match,arr(),i : i = 0
    Set matches = Easp.Str.Match(s,rule)
    ReDim arr(matches.Count-1)
    For Each match In matches
      arr(i) = match.Value
      i = i + 1
    Next
    Set matches = Nothing
    SearchString = arr
  End Function
  
  '在返回HTML中按标签查找字符串
  Public Function Cut(ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
  'tagStart - 要截取的部分的开头
  'tagEnd   - 要截取的部分的结尾
  'tagSelf  - 结果是否包括tagStart和tagEnd
  '           (0或空:不包括,1:包括,2:只包括tagStart,3:只包括tagEnd)
    Cut = CutString(s_ohtml,tagStart,tagEnd,tagSelf)
  End Function
  '在字符串中按标签查找子字符串
  Public Function CutString(ByVal s, ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
    Dim posA, posB, first, between
    posA = instr(1,s,tagStart,1)
    If posA=0 Then
      CutString = ""
      If Easp.Debug Then
        Easp.Error.FunctionName = "Http.CutString"
        Easp.Error.Detail = tagStart
        Easp.Error.Raise "error-http-wrongstart"
      End If
      Exit Function
    End If
    posB = instr(PosA+Len(tagStart),s,tagEnd,1) 
    If posB=0 Then
      CutString = ""
      If Easp.Debug Then
        Easp.Error.FunctionName = "Http.CutString"
        Easp.Error.Detail = tagEnd
        Easp.Error.Raise "error-http-wrongend"
      End If
      Exit Function
    End If
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
    CutString = Mid(s,first,between)
  End Function
  
  '保存返回HTML中的远程图片到本地
  '返回替换图片地址为本地路径后的html代码
  Public Function SaveImgTo(ByVal p)
    SaveImgTo = SaveStringImgTo(s_ohtml,p)
  End Function
  '保存HTML片段中的远程图片到本地
  Public Function SaveStringImgTo(ByVal s, ByVal p)
    Dim a,b, path, i, img, ht, tmp, src
    path = Easp.Str.GetName(s_url, "?")
    path = Left(path, InStrRev(path,"/"))
    '取得图片地址
    a = GetImg(s, False)
    'Easp.Console GetImg(s, False)
    '取得图片标签
    b = GetImg(s, True)
    If Easp.Has(a) Then
      For i = 0 To Ubound(a)
        If Easp.Has(a(i)) Then
          If SaveRandom Then
            'img = Easp.Date.Format(Now,"ymmddhhiiss"&Easp.RandStr("5:0123456789")) & Mid(a(i),InstrRev(a(i),"."))
            If Instr(a(i),".")>0 Then
              img = Easp.NewID() & Mid(a(i),InstrRev(a(i),"."))
            Else
              img = Easp.NewID() & ".jpg"
            End If
          Else
            img = Mid(a(i),InstrRev(a(i),"/")+1)
          End If
          Set ht = Easp.Http.New
          On Error Resume Next
          ht.Get "UTF-8>" & TransPath(s_url, a(i))
          If Err.Number = 0 Then
            tmp = Easp.Fso.SaveAs(p & img, ht.Body)
          End If
          Set ht = Nothing
          If tmp Then
            src = Easp.Str.ReplacePart(b(i), "<img[^>]*?\s+src\s*=\s*((?:"")([^""]+)(?:"")|(?:')([^']+)(?:')|([^\s>]+))[^>]*>", "$1", """" & p & img & """")
            s = Replace(s, b(i), src)
          End If
        End If
      Next
    End If
    SaveStringImgTo = s
  End Function

  '取出html片段中<img>标签，返回数组
  Public Function GetImg(ByVal string, ByVal hasTag)
    Dim s_rule, a, Matches, match, i, s_img, s_src, s_path
    '去掉script标签，因为其中可能含有不正确有图片地址
    string = Easp.Str.Replace(string, "<script([\s\S]+?)</script>", "")
    '匹配img标签的正则
    s_rule = "<img[^>]*?\s+src\s*=\s*((?:"")([^""]+)(?:"")|(?:')([^']+)(?:')|([^\s>]+))[^>]*>"
    i = 0
    If Easp.Str.Test(string, s_rule) Then
      '取消所有的换行和缩进
      string = Replace(string, vbCrLf, " ")
      string = Replace(string, vbTab, " ")
      '正则匹配所有的img标签
      Set Matches = Easp.Str.Match(string, s_rule)
      'Easp.Console Matches
      ReDim a(Matches.Count-1)
      '取出每个img标签
      For Each match In Matches
        '取出图片标签
        s_img = match.Value
        '取出图片地址
        s_src = Replace(Replace(match.SubMatches(0), """", ""), "'", "")
        '更新标签中的src地址
        s_img = Easp.Str.ReplacePart(s_img, s_rule, "$1", """" & s_src & """")
        a(i) = Easp.IIF(hasTag, s_img, s_src)
        i = i + 1
      Next
    Else
      a = Array()
    End If
    GetImg = a
  End Function

  '启用Ajax代理
  Public Sub AjaxAgent()
    Easp.NoCache()
    Dim u, qs, qskey, qf, qfkey, m
    '取得目标地址
    u = Easp.Get("easpurl")
    If Easp.IsN(u) Then Easp.PrintEnd "error:Invalid URL"
    If Instr(u,"?")>0 Then
      qs = "&" & Easp.Str.GetValue(u,"?")
      u = Easp.Str.GetName(u,"?")
    End If
    '传url参数
    If Request.QueryString()<>"" Then
      For Each qskey In Request.QueryString
        If qskey<>"easpurl" Then qs = qs & "&" & qskey & "=" & Request.QueryString(qskey)
      Next
    End If
    u = u & Easp.IfThen(Easp.Has(qs), "?" & Mid(qs,2))
    '如果是Post则同时传Form数据
    m = Request.ServerVariables("REQUEST_METHOD")
    If m = "POST" Then
      If Request.Form()<>"" Then
        For Each qfkey In Request.Form
          qf = qf & "&" & qfkey & "=" & Request.Form(qfkey)
        Next
        Data = Mid(qf,2)
      End If
      Easp.PrintEnd Post(u)
    Else
      Easp.PrintEnd [Get](u)
    End If
  End Sub
  
  '将目录路径转换为目标页面的绝对路径
  '参数：  url - 目标页面，path将以此url为基准
  '       path - 要转换的目录
  '示例： TransPath("http://www.easyaps.cn/test/mypage.html", "/path1/page2.jpg")
  '       返回： http://www.easyaps.cn/path1/page2.jpg
  '      TransPath("http://www.easyaps.cn/test/mypage.html", "path1/page2.jpg")
  '       返回： http://www.easyaps.cn/test/path1/page2.jpg
  Private Function TransPath(ByVal url, ByVal path)
    '如果本来就是绝对路径则直接取出
    If Left(path,7)="http://" Or Left(path,8)="https://" Then TransPath = path : Exit Function
    Dim tmp, ser, fol
    '页面地址
    tmp = Easp.Str.GetName(url,"?")
    '服务器地址
    If Left(url,7)<>"http://" And Left(url,8)<>"https://" Then
      ser = ""
    Else
      ser = Easp.Str.Replace(tmp,"^(https?://[a-zA-Z0-9-.]+)/(.+)$","$1")
    End If
    '页面所在路径
    fol = Mid(tmp,1,InstrRev(tmp,"/"))
    TransPath = Easp.IIF(Left(path,1) = "/", ser, fol) & path
  End Function
  
  'url参数化
  Private Function Serialize__(ByVal a)
    Dim tmp, i, n, v : tmp = ""
    If Easp.IsN(a) Then Exit Function
    If isArray(a) Then
      For i = 0 To Ubound(a)
        n = Easp.Str.GetName(a(i),":")
        v = Easp.Str.GetValue(a(i),":")
        tmp = tmp & "&" & n & "=" & v
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