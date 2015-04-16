<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Use "Http"
Dim http, tmp, rule, arr, i

''=========================
''Demo 7 - 保存远程图片：
Easp.Http.Get "http://www.cnbeta.com/articles/280317.htm"
tmp = Easp.Http.SaveImgTo_(Easp.Http.Html, "imgatlocal/")
Easp.WN Easp.HtmlEncode(tmp)
''=========================

Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>