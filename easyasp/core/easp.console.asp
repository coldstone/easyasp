<%
'######################################################################
'## Easp.Console.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Console Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-04-04
'## Description :   Input and output messages with a console page(window).
'##
'######################################################################

Class EasyASP_Console

  Public Enable, ShowSql, ShowSqlTime
  Public Name, MaxCacheSize, MaxLogSize
  Private s_token
  
  Private Sub Class_Initialize()
    '是否开启控制台
    Enable       = False
    '是否在控制台中自动显示执行的SQL语句
    ShowSql      = True
    '是否在控制台中自动显示执行的SQL语句的执行时间
    ShowSqlTime  = True
    '名称（在一个App中启用多个控制台时以此区分）
    Name         = "easp_console_text"
    '控制台中缓存的内容最大字数
    MaxCacheSize = 8000
    '单条控制台输出的内容最大字数
    MaxLogSize   = 3000
    '防止未授权用户查看Console的token值
    s_token      = ""
  End Sub

  '设置防止未授权访问标识码
  Public Property Let Token(ByVal string)
    s_token = string
  End Property
  
  '写入控制台日志信息
  Public Default Sub Log(ByVal message)
    Dim s_tmp, string
    If Enable Then
      '将输入的对象转为文本
      string = Easp.Str.ToString(message)
      '如果信息长度超过设置的单条信息最大值则进行截取
      If Len(string) > MaxLogSize Then string = Left(string, MaxLogSize)
      '取出已缓存的信息并合并当前信息
      s_tmp = Session(Name) & "> " & string & VbCrLf
      '如果总长度超过设置的缓存内容最大值则进行截取
      If Len(s_tmp) > MaxCacheSize Then s_tmp = Right(s_tmp, MaxCacheSize)
      '放入缓存
      Session(Name) = s_tmp
    End If
  End Sub
  '输出控制台日志信息（在Ajax的目标url中使用）
  Public Sub Out()
    If Not Enable Then Exit Sub
    Dim s_tmp
    '检查token值是否正确
    If Easp.Has(s_token) And Not Easp.Str.IsEqual(Easp.Get("token"), s_token) Then
      Easp.Print "invalidToken"
    Else
      '如果token正确则输出缓存内容到控制台
      Easp.Print Escape(Easp.Str.HtmlEncode(Session(Name)))
      Session(Name) = Empty
    End If
  End Sub
End Class
%>