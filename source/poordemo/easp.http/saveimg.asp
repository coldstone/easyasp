<!--#include file="../../easyasp/easp.asp" --><%
Dim tmp

''=========================
''Demo 7 - 保存远程图片：
Easp.Http.Get "http://www.cnbeta.com/articles/280317.htm"
tmp = Easp.Http.SaveImgTo("imgatlocal/")
Easp.Println Easp.Str.HtmlEncode(tmp)
''=========================
  
Easp.Println ""
Easp.Println "------------------------------------"
Easp.Print "页面执行时间： " & Easp.GetScriptTime & " 秒"
%>