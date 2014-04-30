<!--#include file="../../../easyasp/easp.asp" -->
<%
dim File
Easp.Upload.AllowFileTypes = "*.*"
Easp.Upload.AllowMaxFileSize = "10MB"
Easp.Upload.AllowMaxSize = "20mb"
Easp.Upload.CharSet = "utf-8"
if not Easp.Upload.GetData() then 
	Easp.Println Easp.Upload.Description
else
	Easp.Upload.SavePath = "/_upload"
	Easp.Println "form1 => " & Easp.Post("form1")
	Set File = Easp.Upload.Save("file1",0,true)
	if File.Succeed then
		Easp.Println "文件'" & File.LocalName & "'上传成功，保存位置'" & File.Path & File.FileName & "',文件大小" & File.Size & "字节"
	else
		Easp.Println File.Exception & "<br />"
	end if
end if
%>