<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Debug = False
'EasyASP 上传类 Demo，此例子为UTF-8编码，客户端进度条展示采用了jQuery，需要jQuery库支持
'(此例子是Upload类的基础功能，Upload类尚有功能在开发中，不过目前已有功能基本不会再做更改，可放心使用。)
'友情提示：因为上传类的表单验证是要等文件上传完成之后才能做出判断，所以建议表单的验证请尽量同时做到客户端的验证，否则用户体验将打大折扣。运行此Demo请使用发行版的Easp程序，如果使用开发版，则可能出现进度条临时文件无法正常删除的情况。
Response.Charset = "UTF-8"
'=====================================
Dim random, jsonFile, i, j, f, e
'载入核心类
Easp.Use "Upload"
'使用无组件进度条（默认为不使用）
Easp.Upload.UseProgress = True
'保存进度条数据临时文件的目录（默认为/__uptemp）
'Easp.Upload.ProgressPath = "/__uptemp"
'仅允许上传文件类型(建议先在客户端判断)
Easp.Upload.Allowed = "exe|jpg|gif|png|rar|zip"
'禁止上传的文件类型，如果设置了仅允许上传文件类型，则此设置不生效(建议先在客户端判断)
'Easp.Upload.Denied = "exe|msi|bat|cmd|asp|asa"
'单个文件最大允许值，单位为KB（如果是图片建议在客户端判断）
Easp.Upload.FileMaxSize = 1024*1000
'全部文件最大允许值，单位为KB
Easp.Upload.TotalMaxSize = 1024*3000
'点击上传后执行
If Easp.Get("act") = "upload" Then
	'上传文件保存路径:
	'比如：
	'Easp.Upload.SavePath = "/userFiles/"
	'或者，可用<>带日期标志（参见Easp.DateTime）按日期建立相应文件夹：
	Easp.Upload.SavePath = "uploadfiles/<yyyy>/<mm>/"
	'保存时使用随机文件名，默认为False，即不使用
	Easp.Upload.Random = True
	'是否自动建立不存在的文件夹，默认为True，即会自动建立
	'Easp.Upload.AutoMD = False
	'获取上传的唯一KEY用于生成进度条数据Json文件给js调用
	Easp.Upload.Key = Easp.Get("json")
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
	Easp.WC "<ul>"
	'取表单项的值，仍然用Easp.Post^_^，还有一种用法是Easp.Upload.Form("nick")，但是就没有Post的多参数啦：
	Easp.WC "<li>用户名 : " & Easp.Post("nick:s") & "</li>"
	'表单项是集合，可以遍历：
	For Each i In Easp.Upload.Form
		Easp.WC "<li>表单项 '" & i & "' 的值 : " & Easp.Upload.Form(i) & "</li>"
	Next
	Easp.WC "</ul>"
	Easp.WN "共 " & Easp.Upload.File.Count & " 个文件上传框，本次上传成功了 " & Easp.Upload.Count & " 个文件。成功上传的文件信息如下："
	Easp.WN "================="
	'所有的上传文件也是一个集合，同样可以遍历：
	For Each j In Easp.Upload.File
		Set f = Easp.Upload.File(j)
		'只列出上传的文件
		If f.Size>0 Then
			Easp.WN "文件原位置：" & f.Client
			Easp.WN "文件原目录：" & f.OldPath
			Easp.WN "文件大小：" & f.Size
			Easp.WN "文件名称：" & f.Name
			Easp.WN "文件扩展名：" & f.Ext
			Easp.WN "文件MIME类型：" & f.MIME
			Easp.WN "新路径："& f.NewPath
			Easp.WN "新名称："& f.NewName
			Easp.WN "Web路径：" & f.WebPath & Server.URLEncode(f.NewName)
			Easp.WN "================="
		End If
	Next
	'显式的释放，才能删除进度条数据文件
	Set Easp.Upload = Nothing
	Easp.WE "<a href=""../upload/"">继续上传</a>"
End If
'生成本次上传的唯一KEY
random = Easp.Upload.GenKey
'获取给js使用的Json文件的地址
jsonFile = Easp.Upload.ProgressFile(random)
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
.upload .progress, .upload #progress { font-size:12px;position:absolute;top:80px;left:256px;height:130px;width:240px; background-color:#E4E4E4;}
.upload .progress { line-height:130px; text-align:center;}
.upload #progress {display:none; color:#036;}
.upload #progress .info{padding:10px 0;text-align:center; line-height:1.5em;}
.upload .progress-bar{border:1px solid #069; position:absolute; height:14px; width:200px;left:20px; bottom:10px;}
.upload #uploadPercent{position:absolute;top:0;left:0;width:100%;font-size:11px;text-align:center;}
.upload #progressBar{position:absolute;top:0;left:0;width:0; height:14px;background-color:#0099E3;}
-->
</style>
</head>
<body>
<fieldset class="upload"><legend>Easp.Upload上传文件示例 （多文件带进度条）</legend>
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
    <div id="progress">
        <div class="info">
			<strong>正在上传，请稍候…</strong><br />
            总大小：<span id="uploadTotalSize">0 KB</span> / 已上传：<span id="uploadSize">0 KB</span><br />
			上传速度：<span id="uploadSpeed">0 KB/S</span><br />
            总共时间：<span id="uploadTotalTime">00:00:00</span><br />
            剩余时间：<span id="uploadRemainTime">00:00:00</span>
        </div>
		<div class="progress-bar">
			<div id="progressBar"></div>
			<div id="uploadPercent">0%</div>
		</div>
    </div>
</form>
</fieldset>
</body>
<script language="javascript">
$('#formUpload').submit(function(){
	//$('#progress').show();
	//return false;
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
		this.action = 'index.asp?act=upload&json=<%=random%>';
		//显示进度条
		startProgress('<%=jsonFile%>');
		return true;
	}
});

//显示进度条
function startProgress(path){
	//0.5秒后启动进度条
	setTimeout('readProgress("' + path + '")',500);
}
//读取进度条
function readProgress(path){
	var percent = 0;
	try{
		//Ajax读取
		$.get(path,{rnd:Math.floor(Math.random()*10000)},function(d){
			//解析Json
			var progress = eval('('+d+')');
			//已上传大小 和 总大小
			$('#uploadTotalSize').text(progress.total);
			$('#uploadSize').text(progress.uploaded);
			//上传速度
			$('#uploadSpeed').text(progress.speed);
			//上传总共需要时间（估计值）
			$('#uploadTotalTime').text(progress.totaltime);
			//上传剩余时间
			$('#uploadRemainTime').text(progress.remaintime);
			//上传百分比，更新显示状态
			percent = progress.percent;
			$('#uploadPercent').text(percent+'%');
			$('#progressBar').css({'width':percent*2+'px'});
		});
	} catch(e){ }
	//上传如未完成继续刷新(时间为0.5秒刷新一次)
	if (percent<100){
		setTimeout('readProgress("'+path+'")',500);
		//显示进度条
		$('#progress').show();
	} else {
		$('.progress').text('上传成功！正在保存文件…');
		$('#progress').hide();
	}
}
</script>
</html>
<% 
'释放Easp对象
Set Easp = Nothing
'做功能不难，但要做例子把所有功能展现出来真难啊，大家请多给意见！
%>