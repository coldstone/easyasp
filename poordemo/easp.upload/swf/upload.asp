<!--#include file="../../../easyasp/easp.asp" -->
<%
Dim File
Easp.Upload.AllowMaxSize="200mb"
Easp.Upload.AllowMaxFileSize="200mb"
Easp.Upload.AllowFileTypes="*.*" 
Easp.Upload.Charset="utf-8"
if not Easp.Upload.GetData() then
	Easp.Echo "{err:true,msg:'" & Easp.Upload.Description & "'}"
else
	Easp.Upload.SavePath = "/_upload"
	set File=Easp.Upload.files("filedata") 
	if Easp.Upload.Save(File,0,true).Succeed then
		Easp.Echo "{err:false,msg:'upload',name:'" & File.filename & "',src:'" & File.LocalName & "',name2:'" & Easp.Upload.Post("name") & "'}"
	else
		Easp.Echo "{err:true,msg:'" & File.Exception & "'}"
	end if
	set File=nothing
end if
%>
