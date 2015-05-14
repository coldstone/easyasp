<!--#include file="../../../easyasp/easp.asp" -->
<%
dim File
Easp.Var("test1") = "test1"
Easp.Upload.AllowFileTypes = "*.jpg"
Easp.Upload.AllowMaxFileSize = "10MB"
Easp.Upload.AllowMaxSize = "20mb"
Easp.Upload.CharSet = "utf-8"
Easp.Println "Easp.Var(""form1"") => " & Easp.Var("form1")  '在调用 Easp.Upload.GetData() 之前是取不到表单数据的
Easp.Println "Easp.Var(""act"") => " & Easp.Var("act") 'querystring的值则可以随时调用
Easp.Println "Easp.Var(""test1"") => " & Easp.Var("test1")
if not Easp.Upload.GetData() then 
	Easp.Println Easp.Upload.Description
else
  Easp.Var("test2") = "test2"
	Easp.Upload.SavePath = "/_upload"
	Easp.Println "Easp.Var(""test2"") => " & Easp.Var("test2")
	Easp.Println "Easp.Post(""form1"") => " & Easp.Post("form1")
	Easp.Println "Easp.Db.ToSql(""delete from T where Tname in ({(form1)})"") =>" & _
	             Easp.Db.ToSql("delete from T where Tname in ({(form1)})")
	Set File = Easp.Upload.Save("file1",0,true)
	if File.Succeed then
		Easp.Println "文件'" & File.LocalName & "'上传成功，保存位置'" & File.Path & File.FileName & "',文件大小" & File.Size & "字节"
	else
		Easp.Println File.Exception & "<br />"
	end if
end if
%>