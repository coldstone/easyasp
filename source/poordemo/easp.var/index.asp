<!--#include file="../../easyasp/easp.asp" --><%
Easp.NoCache()
Easp.Var("myname") = "Lin"
If Easp.Has(Request.Form) Then
  Easp.println "Easp.Var(""easp.newid"") : " & Easp.Var("easp.newid")
  Easp.println "Easp.Var(""url"") : " & Easp.Var("url")
  Easp.println "Easp.Var(""myname"") : " & Easp.Var("myname")
  Easp.println "Easp.Var(""get.username"") : " & Easp.Var("get.username")
  Easp.println "Easp.Var(""post.username"") : " & Easp.Var("post.username")
  Easp.println "Easp.Var(""username"") : " & Easp.Var("username")
  Easp.println "Easp.Var(""msg"") : " & Easp.Var("msg")
  Easp.println "Easp.Var(""action"") : " & Easp.Var("action")
  If Easp.Var.Has("action_array") Then
    'Easp.print "如果同一名称URL参数有多个值：Request.QueryString(""action"").Count : "
    'Easp.println Request.QueryString("action").Count
    Easp.println "Easp.Var(""action_array"") : " & Easp.Str.ToString(Easp.Var("action_array"))
  End If
  Easp.println "Easp.Var(""type"") : " & Easp.Var("type")
  If Easp.Var.Has("type_array") Then
    'Easp.print "如果同一名称表单有多个值：Request.Form(""type"").Count : "
    'Easp.println Request.Form("type").Count
    Easp.println "Easp.Var(""type_array"") : " & Easp.Str.ToString(Easp.Var("type_array"))
  End If
  Easp.Println "Easp.Var(""server.remote_addr"") : " & Easp.Var("server.remote_addr")
  Easp.Println "Easp.Var(""server.http_user_agent"") : " & Easp.Var("server.http_user_agent")
  '显示所有的变量
  Easp.println "=============================="
  Easp.println "遍历所有的EasyAsp变量："
  Dim vars, key 
  Set vars = Easp.Var.GetObject()
  For Each key In vars
    Easp.print "Easp.Var(""" & key & """) : "
    Easp.println Easp.Str.ToString(vars(key))
  Next
  Set vars = Nothing
End If
%>
<form action="?action=save&username=coldstone&action=update" method="post">
  username: <input type="text" size="60" name="username" value="ray" /><br />
  msg: <input type="text" size="60" name="msg" value="I'm here" /><br />
  <input type="checkbox" name="type" value="1" checked="checked" />type1
  <input type="checkbox" name="type" value="2" checked="checked" />type2<br />
  <button type="submit">Submit to "?action=save&username=coldstone&action=update"</button>
</form>