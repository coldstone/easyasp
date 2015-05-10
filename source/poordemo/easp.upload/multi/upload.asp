<!--#include file="../../../easyasp/easp.asp" -->
<%
Dim File,F
Easp.Upload.AllowFileTypes = "jpg|png|gif"
Easp.Upload.AllowMaxFileSize = "1MB"
Easp.Upload.AllowMaxSize = "20mb"
Easp.Upload.CharSet = "utf-8"
if not Easp.Upload.GetData() then 
	Easp.PrintEnd Easp.Upload.Description
else
	Easp.Upload.SavePath = "/_upload"
	Easp.Println "<b>保存所有文件： </b>"
	for each file in Easp.Upload.Files("-1")
		Set F = Easp.Upload.Save(file,0,true)
		if F.Succeed then
			Easp.Println "文件'" & F.LocalName & "'上传成功，保存位置'" & F.Path & F.filename & "',文件大小" & F.size & "字节"
		else
			Easp.Println F.Exception
		end if		
	next
end if
%>