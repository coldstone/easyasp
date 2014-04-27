<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.db.Conn = Easp.db.OpenConn(0,"EasyASP","sa:jpzx_1860@192.168.133.2")
Easp.Debug = False
'EasyASP 上传类 Demo，此例子为UTF-8编码，客户端进度条展示采用了jQuery，需要jQuery库支持
'(此例子是Upload类的基础功能，Upload类尚有功能在开发中，不过目前已有功能基本不会再做更改，可放心使用。)
'友情提示：因为上传类的表单验证是要等文件上传完成之后才能做出判断，所以建议表单的验证请尽量同时做到客户端的验证，否则用户体验将打大折扣。运行此Demo请使用发行版的Easp程序，如果使用开发版，则可能出现进度条临时文件无法正常删除的情况。
Response.Charset = "UTF-8"
'=====================================
Dim i, j, f, e
'载入核心类
Easp.Use "Upload"
'仅允许上传文件类型(建议先在客户端判断)
Easp.Upload.Allowed = "exe|jpg|gif|png|rar|zip"
'单个文件最大允许值，单位为KB（如果是图片建议在客户端判断）
Easp.Upload.FileMaxSize = 1024*10
'全部文件最大允许值，单位为KB
Easp.Upload.TotalMaxSize = 1024*30
'点击上传后执行
If Easp.Get("act") = "upload" Then
	'上传文件保存路径:
	'比如：
	'Easp.Upload.SavePath = "/userFiles/"
	'或者，可用<>带日期标志（参见Easp.DateTime）按日期建立相应文件夹：
	Easp.Upload.SavePath = "uploadfiles/<yyyy>/<mm>/"
	'重要方法：开始上传
	Easp.Upload.StartUpload()
	'捕捉错误信息并弹出警告窗口返回上传页
	If Easp.Error.LastError>"" Then
		'如果出错就释放对象，同时删除出错文件的进度条数据文件
		Set Easp.Upload = Nothing
		Easp.Alert Easp.Error(Easp.Error.LastError)
	End If
	'保存全部上传文件
	Easp.Upload.SaveAll
	'或者保存单个文件：
	'Easp.Upload.File("file1").Save
	'或者保存单个文件到具体路径(另存为)：
	'Easp.Upload.File("file1").SaveAs Easp.Upload.File("file1").NewPath & "这是新文件名." & Easp.Upload.File("file1").Ext
	'显式的释放，才能删除进度条数据文件
	Set Easp.Upload = Nothing
	Easp.WE "<a href=""../upload/"">继续上传</a>"
End If
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>EasyAsp Upload Demo</title>
<script type="text/javascript" src="../jquery-1.4.1.min.js"></script>
<style type="text/css">
<!--
/*表单样式*/
.upload {font-size:14px; font-family:Tahoma;width:500px; padding:20px 20px;}
.upload p{ padding:0; margin:0 0 10px 0;}
.upload input{ font-size:12px;font-family:Tahoma; padding:4px;}
.upload input.ipt{ width:436px;}
.upload input.ipts{ width:180px;}
.upload .btns{ padding:10px 0;}
/*进度条样式*/
.upload #formUpload{position:relative;}
.upload .progress { font-size:12px;position:absolute;top:80px;left:256px;height:130px;width:240px; background-color:#E4E4E4;}
.upload .progress { line-height:130px; text-align:center;}
-->
</style>
</head>
<body>
<fieldset class="upload"><legend>Easp.Upload上传文件示例 （同时把地址保存到数据库）</legend>
<!-- form的method是post，enctype是multipart/form-data，这在上传表单里是必需的 -->
<form id="formUpload" method="post" enctype="multipart/form-data">
	<p>昵 称：<input type="text" name="nick" class="ipt" /></p>
	<p>密 码：<input type="password" name="pwd" class="ipt" /></p>
	<p>附件1：<input name="file1" type="file" /></p>
	<p>附件2：<input name="file2" type="file" /></p>
	<p>附件3：<input name="file3" type="file" /></p>
	<p>附件3：<input name="file4" type="file" /></p>
	<p>允许上传类型：<%= Replace(Easp.Upload.Allowed,"|", ", ") %></p>
	<p>单个文件最大允许值：<%= Easp.Upload.FileMaxSize %> KB</p>
	<p>全部文件最大允许值：<%= Easp.Upload.TotalMaxSize %> KB</p>
	<div class="btns"><input type="submit" id="btnSubmit" value="确认提交"/></div>
	<div class="progress">请选择一个或多个要上传的文件！</div>
  <%
	Easp.Upload.Conn = Easp.db.Conn
	Easp.wn TypeName(Easp.Upload.Conn)
	%>
</form>
</fieldset>
</body>
<script language="javascript">
$('#formUpload').submit(function(){
	var flag = false;
	$(this).find(':file').each(function(){
		if ($(this).val()!=''){
			flag = true;
			return false;
		}
	});
	if (!flag) {
		alert('请至少上传一个文件！');
		return false;
	} else {
		//在Form的action中加入上传的唯一KEY
		this.action = '?act=upload';
		return true;
	}
});
</script>
</html>
<% 
'释放Easp对象
Set Easp = Nothing
%>