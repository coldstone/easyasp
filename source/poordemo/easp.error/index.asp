<!--#include file="../../easyasp/easp.asp" --><%
'Easp.Debug = False
'Easp.Error.Redirect = False
Easp.Error.OnErrorContinue = True
Easp.Error.ConsoleDetail = False
Easp.Println Link("", Easp.GetUrl(-3) & "?" & Easp.NewID, "")
Dim s
'Easp.Console "[Error]数据库读取错误。"
'Easp.Console Easp.Error.Debug

'For Each s In Request.ServerVariables
'  Easp.Println s & " : " & Request.ServerVariables(s)
'Next
On Error Resume Next
'Easp.Db.SetConn 0, "Easp", "sa:pass@(local))"
Easp.Ext("check").Meinv
Dim conn
'Set conn = Easp.Db.GetConn()
'Err.Raise 45, "my error"
'Easp.Console Easp.Error.Redirect
'Easp.Error.Detail = "(sa:pass@(local))"
'Easp.Error.Raise 12

Function Link(ByVal string, ByVal url, ByVal attr)
  Dim a
  a = Easp.IfHas(string, url)
  Link = "<a href=""" & url & """" & Easp.IfThen(Easp.Has(attr), " " & attr) & ">" & a & "</a>"
End Function

Easp.Println "============================"
Easp.Println Easp.GetScriptTime & "s"

%>
