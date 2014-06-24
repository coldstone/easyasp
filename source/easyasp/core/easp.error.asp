<%
'######################################################################
'## easp.error.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Exception Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-04-19 02:14:28
'## Description :   Deal with the EasyASP Exception
'##
'######################################################################
Class EasyASP_Error
  Private b_redirect, b_continue, b_console
  Private i_errNum, i_delay
  Private s_title, s_url, s_css, s_msg, s_funName
  Private o_err, a_detail
  Private e_err, e_conn, e_dom
  Private Sub Class_Initialize()
    i_errNum    = ""
    i_delay     = 5
    s_title     = Easp.Lang("error-title")
    b_redirect  = False
    b_console   = True
    b_continue  = False
    s_url       = "javascript:history.go(-1)"
    Set o_err   = Server.CreateObject("Scripting.Dictionary")
    o_err.CompareMode = 1
  End Sub
  Private Sub Class_Terminate()
    Set o_err = Nothing
    If IsObject(e_err) Then Set e_err = Nothing
    If IsObject(e_conn) Then Set e_conn = Nothing
    If IsObject(e_dom) Then Set e_dom = Nothing
  End Sub
  '设置或读取自定义的错误代码和错误信息
  Public Default Property Get E(ByVal n)
    If IsNumeric(n) Then n = "E" & n
    If o_err.Exists(n) Then
      E = Join(o_err(n), "|")
    Else
      E = Easp.Lang("error-unkown")
    End If
  End Property
  Public Property Let E(ByVal n, ByVal s)
    Dim a_info, i_tmp
    If Easp.Has(n) And Easp.Has(s) Then
      If IsNumeric(n) Then n =  "E" & n
      a_info = Split(s, "|")
      i_tmp = UBound(a_info)
      If i_tmp < 2 Then
        a_info = Split(s & String(2 - i_tmp, "|"), "|")
      End If
      o_err(n) = a_info
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
  '设置显示错误信息时的详细信息替换参数
  Public Property Let Detail(ByVal arr)
    a_detail = arr
  End Property
  '设置和读取出错函数名
  Public Property Get FunctionName()
    FunctionName = s_funName
  End Property
  Public Property Let FunctionName(ByVal string)
    s_funName = string
  End Property
  '设置和读取页面是否自动转向
  Public Property Get [Redirect]
    [Redirect] = b_redirect
  End Property
  Public Property Let [Redirect](ByVal b)
    b_redirect = b
  End Property
  '设置和读取Debug模式下出错后是否继续运行后面的代码
  '说明：普通模式下总是继续运行
  Public Property Get OnErrorContinue
    OnErrorContinue = b_continue
  End Property
  Public Property Let OnErrorContinue(ByVal bool)
    b_continue = bool
  End Property
  '设置和读取是否在控制台内显示详细错误信息
  Public Property Get ConsoleDetail
    ConsoleDetail = b_console
  End Property
  Public Property Let ConsoleDetail(ByVal bool)
    b_console = bool
  End Property
  '设置和读取发生错误后的跳转页地址
  '说明：如不设置此属性，则默认为返回前一页
  Public Property Get Url
    Url = s_url
  End Property
  Public Property Let Url(ByVal s)
    s_url = s
  End Property
  '设置和读取自动跳转页面等待时间（秒）
  Public Property Get Delay
    Delay = i_delay
  End Property
  Public Property Let Delay(ByVal i)
    i_delay = i
  End Property
  '设置和读取显示错误信息DIV的CSS样式名称
  Public Property Get ClassName
    ClassName = s_css
  End Property
  Public Property Let ClassName(ByVal s)
    s_css = s
  End Property

  'Dom和Connection错误
  Public Sub SetErrors(ByRef e, ByRef ec, ByRef ed)
    If isObject(e) Then Set e_err = e
    If IsObject(ec) Then Set e_conn = ec
    If isObject(ed) Then Set e_dom = ed
  End Sub
  
  '生成一个错误(常用于开发者错误模式)
  Public Sub Raise(ByVal n)
    If Easp.isN(n) Then Exit Sub
    If IsNumeric(n) Then n = "E" & n
    If Not IsObject(e_err) Then Set e_err = Err
    Dim b_consoleDetail, b_isEnd
    Dim msg
    '得到已定义错误信息
    msg = o_err(n)
    '如果是Debug模式，出错后是否继续运行
    b_isEnd = Easp.IIF(Easp.Debug , Not b_continue, False)
    '在控制台内输出错误信息
    InConsole msg, b_console
    If b_isEnd Then
      Easp.PrintEnd ShowErrorMsg(msg)
    Else
      Easp.Print ShowErrorMsg(msg)
    End If
    i_errNum = n
    s_msg = ""
  End Sub
  
  '立即抛出一个错误信息(常用于用户错误模式)
  Public Sub Throw(ByVal msg)
    Dim a_info, i_tmp
    If Easp.Has(msg) Then
      a_info = Split(msg, "|")
      i_tmp = UBound(a_info)
      If i_tmp < 2 Then
        a_info = Split(msg & String("|", 2 - i_tmp), "|")
      End If
      Easp.Print ShowErrorMsg(a_info)
    End If
  End Sub
  '在控制台中抛出错误信息
  Public Sub Console(ByVal n)
    If Easp.isN(n) Then Exit Sub
    If IsNumeric(n) Then n = "E" & n
    Dim msg
    msg = o_err(n)
    InConsole msg, Easp.Debug
  End Sub
  
  '控制台输出错误：
  Private Sub InConsole(ByVal msg, ByVal hasDetail)
    If Easp.Console.Enable Then
      Dim SB : Set SB = Easp.Str.StringBuilder()
      SB.Append "[Error] "
      SB.Append msg(0)
      If hasDetail Then
        SB.Append " ("
        If Easp.Has(msg(1)) Then
          If Left(msg(1), 1) = ":" Then msg(1) = Mid(msg(1), 2)
          SB.Append "详细信息：" & Easp.Str.Format(msg(1), a_detail) & "; "
        End If
        If Easp.Has(s_funName) Then SB.Append "来源函数：" & s_funName & "; "
        SB.Append "请求URL：" & Easp.GetUrl("") & "; "
        SB.Append "请求方式：" & Request.ServerVariables("REQUEST_METHOD") & "; "
        Dim s_ref : s_ref = Request.ServerVariables("HTTP_REFERER")
        If Easp.Has(s_ref) Then
          SB.Append "来源URL：" & s_ref
        End If
        If Err.Number <> 0 Then
          SB.Append "; 错误代码：" & Err.Number
          SB.Append "; 错误描述：" & Err.Description
          SB.Append "; 错误来源：" & Err.Source
        End If
        If Easp.Has(msg(2)) Then SB.Append"; 处理建议：" & msg(2)
        SB.Append ")"
      End If
      Easp.Console SB.ToString()
      Set SB = Nothing
    End If
  End Sub
  
  '显示错误信息框
  Private Function ShowErrorMsg(ByVal msg)
    Dim SB, key, s_ref
    s_ref = Request.ServerVariables("HTTP_REFERER")
    Set SB = Easp.Str.StringBuilder()
    If Easp.IsN(s_css) Then
      s_css = "easp-error"
      SB.Append "<style>.easp-error{width:70%;font-size:12px;font-family:""Microsoft Yahei"";margin:10px auto;padding:10px 20px;}.easp-error legend{margin:0 0 5px 0;padding:0 10px;font-size:14px;font-weight:bolder;}.easp-error p{margin:0 0 10px 0;padding:0;}.easp-error p.msg{font-size:14px;}.easp-error p a:link{color:#09F;}.easp-error p a:hover{color:#090;}.easp-error h3{font-size:12px;margin:0 0 10px 0;padding:0;font-weight:normal;}.easp-error h3 .title{font-weight:bolder;}.easp-error h3 a{color:#090;text-decoration:none;font-family:consolas;}.easp-error .info{margin-bottom:10px;margin-top:-6px;}.easp-error ul.list{margin:0;padding:0;}.easp-error ul.list li{list-style:none;margin:0 24px;color:#666;line-height:1.6em;word-break:break-all;}.easp-error ul.list li strong{color:#555;}</style>"
    End If
    SB.Append "<fieldset id=""easpError"" class="""
    SB.Append s_css
    SB.Append """>"
    SB.Append "<legend>"
    SB.Append s_title
    SB.Append "</legend>"
    SB.Append "<p class=""msg"">"
    SB.Append msg(0)
    SB.Append "</p>"
    If Easp.Debug Then
      '显示详细错误信息
      SB.Append "<h3><a href=""javascript:toggle('easp_err_detail')"" id=""easp_err_detail_m"">[-]</a> <span class=""title"">详细错误信息</span></h3>"
      SB.Append "<div class=""info"" id=""easp_err_detail"">"
      SB.Append "<ul class=""list"">"
      If Easp.Has(msg(1)) Then
        If Left(msg(1), 1) = ":" Then msg(1) = Mid(msg(1), 2)
        SB.Append "<li><strong>错误信息 : </strong>"
        SB.Append Easp.Str.Format(msg(1), a_detail)
        SB.Append "</li>"
      End If
      If Easp.Has(s_funName) Then
        SB.Append "<li><strong>来源函数 : </strong>"
        SB.Append s_funName
        SB.Append "</li>"
      End If
      SB.Append "<li><strong>请求URL : </strong>"
      SB.Append Easp.GetUrl("")
      SB.Append "</li>"
      SB.Append "<li><strong>请求方式 : </strong>"
      SB.Append Request.ServerVariables("REQUEST_METHOD")
      SB.Append "</li>"
      IF Easp.Has(s_ref) Then
        SB.Append "<li><strong>来源URL : </strong>"
        SB.Append s_ref
        SB.Append "</li>"
      End If
      If IsObject(e_conn) Then
        If e_conn.Errors.Count > 0 Then
          If e_conn.Errors(0).Number <> 0 Then
            With e_conn.Errors(0)
              SB.Append "<li><strong>数据库类型 : </strong>"
              SB.Append Easp.Db.GetTypeVersion(e_conn)
              SB.Append "</li>"
              SB.Append "<li><strong>错误代码 : </strong>"
              SB.Append .Number
              SB.Append "</li>"
              SB.Append "<li><strong>错误描述 : </strong>"
              SB.Append .Description
              SB.Append "</li>"
              SB.Append "<li><strong>源错代码 : </strong>"
              SB.Append .NativeError
              SB.Append "</li>"
              SB.Append "<li><strong>错误来源 : </strong>"
              SB.Append .Source
              SB.Append "</li>"
              SB.Append "<li><strong>SQL 错误码 : </strong>"
              SB.Append .SQLState
              SB.Append "</li>"
            End With
          End If
        End If
      End If
      If IsObject(e_dom) Then
        If e_dom.errorCode <> 0 Then
          With e_dom
            SB.Append "<li><strong>DOM错误代码 : </strong>"
            SB.Append .errorCode
            SB.Append "</li>"
            SB.Append "<li><strong>DOM错误原因 : </strong>"
            SB.Append .reason
            SB.Append "</li>"
            SB.Append "<li><strong>DOM错误来源 : </strong>"
            SB.Append .url
            SB.Append "</li>"
            SB.Append "<li><strong>DOM错误行号 : </strong>"
            SB.Append .line
            SB.Append "</li>"
            SB.Append "<li><strong>DOM错误位置 : </strong>"
            SB.Append .linepos
            SB.Append "</li>"
            SB.Append "<li><strong>DOM源文本 : </strong>"
            SB.Append .srcText
            SB.Append "</li>"
          End With
        End If
      End If
      If Not IsObject(e_conn) And Not IsObject(e_dom) Then
        If e_err.Number <> 0 Then
          SB.Append "<li><strong>错误代码 : </strong>"
          SB.Append e_err.Number
          SB.Append "</li>"
          SB.Append "<li><strong>错误描述 : </strong>"
          SB.Append e_err.Description
          SB.Append "</li>"
          SB.Append "<li><strong>错误来源 : </strong>"
          SB.Append e_err.Source
          SB.Append "</li>"
        End If
      End If

      If Easp.Has(msg(2)) Then
        If Left(msg(2), 1) = ":" Then msg(2) = Mid(msg(2), 2)
        SB.Append "<li><strong>处理建议 : </strong>"
        SB.Append msg(2)
        SB.Append "</li>"
      End If
      SB.Append "</ul>"
      SB.Append "</div>"
      '显示QueryString
      If Request.QueryString.Count > 0 Then
        SB.Append "<h3><a href=""javascript:toggle('easp_err_querystring')"" id=""easp_err_querystring_m"">[-]</a> <span class=""title"">Query String 参数</span></h3>"
        SB.Append "<div class=""info"" id=""easp_err_querystring"">"
        SB.Append "<ul class=""list"">"
        For Each key In Request.QueryString
          SB.Append "<li><strong>"
          SB.Append key
          SB.Append " : </strong>"
          SB.Append Request.QueryString(key)
          SB.Append "</li>"
        Next
        SB.Append "</ul>"
        SB.Append "</div>"
      End If
      '显示Form
      If Request.Form.Count > 0 Then
        SB.Append "<h3><a href=""javascript:toggle('easp_err_form')"" id=""easp_err_form_m"">[-]</a> <span class=""title"">表单数据</span></h3>"
        SB.Append "<div class=""info"" id=""easp_err_form"">"
        SB.Append "<ul class=""list"">"
        For Each key In Request.Form
          SB.Append "<li><strong>"
          SB.Append key
          SB.Append " : </strong>"
          SB.Append Request.Form(key)
          SB.Append "</li>"
        Next
        SB.Append "</ul>"
        SB.Append "</div>"
      End If
      '显示HTTP报头
      Dim i, keyName
      SB.Append "<h3><a href=""javascript:toggle('easp_err_http')"" id=""easp_err_http_m"">[+]</a> <span class=""title"">HTTP 报头</span></h3>"
      SB.Append "<div class=""info"" id=""easp_err_http"" style=""display:none;"">"
      SB.Append "<ul class=""list"">"
      key = Split(Request.ServerVariables("ALL_HTTP"), vbLf)
      For i = 0 To UBound(key)-1
        keyName = LCase(Easp.Str.Replace(Easp.Str.GetColonName(key(i)), "^http_", ""))
        SB.Append "<li><strong>"
        SB.Append UCase(Left(keyName,1)) & Mid(keyName,2)
        SB.Append " : </strong>"
        SB.Append Easp.Str.GetColonValue(key(i))
        SB.Append "</li>"
      Next
      SB.Append "</ul>"
      SB.Append "</div>"
    Else
      '显示普通模式详细错误信息
      If (Easp.Has(msg(1)) And Left(msg(1), 1) <> ":") Or Easp.Has(msg(2)) Then
        SB.Append "<h3><a href=""javascript:toggle('easp_err_detail')"" id=""easp_err_detail_m"">[-]</a> <span class=""title"">详细错误信息</span></h3>"
        SB.Append "<div class=""info"" id=""easp_err_detail"">"
        SB.Append "<ul class=""list"">"
        If Easp.Has(msg(1)) And Left(msg(1), 1) <> ":" Then
          SB.Append "<li><strong>错误信息 : </strong>"
          SB.Append Easp.Str.Format(msg(1), a_detail)
          SB.Append "</li>"
        End If
        If Easp.Has(msg(2)) And Left(msg(2), 1) <> ":" Then
          SB.Append "<li><strong>处理建议 : </strong>"
          SB.Append msg(2)
          SB.Append "</li>"
        End If
        SB.Append "</ul>"
        SB.Append "</div>"
      End If
    End If
    SB.Append "<p class=""redirect"">"
    If Easp.Str.IsSame(s_ref, Easp.GetUrl("")) Or Easp.IsN(s_ref) Then
      b_redirect = False
      s_url = "javascript:location.reload(true)"
    End If
    If b_redirect Then
      SB.Append "页面将在<span id=""easp_timeoff"">"
      SB.Append i_delay
      SB.Append "</span>秒钟后跳转，如果浏览器没有正常跳转，"
    End If
    SB.Append "<a href="""
    SB.Append s_url
    SB.Append """>请点击此处"
    If Easp.Str.IsSame(s_url, "javascript:history.go(-1)") Then
      SB.Append "返回"
    ElseIf Easp.Str.IsSame(s_url, "javascript:location.reload(true)") Then
      SB.Append "刷新"
    Else
      SB.Append "继续"
    End If
    SB.Append "</a></p>"
    SB.Append "<script type=""text/javascript"">function toggle(id){var el = document.getElementById(id);var a = document.getElementById(id+""_m"");if(a.innerHTML==""[-]""){el.style.display = ""none"";a.innerHTML = ""[+]"";}else if(a.innerHTML==""[+]""){el.style.display = """";a.innerHTML = ""[-]"";}}"
    If b_redirect Then
      SB.Append "function timeMinus(){var el = document.getElementById(""easp_timeoff"");var timeLeft = parseInt(el.innerHTML);el.innerHTML = timeLeft - 1;} setInterval(timeMinus, 1000);"
      SB.Append "setTimeout(function(){"
      If Easp.Str.IsSame(Left(s_url,11), "javascript:") Then
        SB.Append Mid(s_url, 12)
      Else
        SB.Append "location.href='"
        SB.Append s_url
        SB.Append "'"
      End If
      SB.Append "},"
      SB.Append i_delay * 1000
      SB.Append ");"
    End If
    SB.Append "</script></fieldset>"
    ShowErrorMsg = SB.ToString()
    Set SB = Nothing
  End Function
  
  '显示已定义的所有错误代码及信息，返回Json格式
  Public Function Defined()
    Defined = Easp.Str.ToString(o_err)
  End Function
End Class
%>