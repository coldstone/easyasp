<!--#include file="../../easyasp/easp.asp" --><!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8" />
  <title>EasyASP 压缩解压演示</title>
  <style>
    body,html{font-size:14px;}
    fieldset {margin:10px;padding:10px;}
    .right {text-align:right;}
    .file-list {width:320px;height:320px;float:left;}
    .file-list select{width:300px;height:300px;font-family:consolas;}
    .form1{width:500px;float:left;}
    .form1 input{width:100%;}
    .form1 textarea{width:100%;height:76px;}
    .form1 div{margin-bottom:5px;}
  </style>
</head>
<body>
  <fieldset>
    <legend>EasyASP 无损压缩(.tar格式)演示</legend>
    <form method="post" id="form1">
      <div class="file-list">
        <select id="files" name="files" multiple="true"></select>
      </div>
      <div class="form1">
        <div>请先从左侧选择要压缩的文件夹和文件（可按住Ctrl或Shift多选）</div>
        <div>在压缩包内生成一个新文件：<br /><input type="text" name="newfile" id="newfile" value="readme.txt" /></div>
        <div>要生成的新文件的内容：<br />
        <textarea name="newfilecontent" id="newfilecontent">备份人：coldstone
网站：<%=Easp.GetUrl(-1)%>
备份时间：<%=Easp.Date.Format(Now(), "y-mm-dd hh:ii:ss")%></textarea></div>
        <div>在压缩包内生成一个空文件夹：<br /><input type="text" name="newfolder" id="newfolder" value="中文文件夹/这是空的/" /></div>
        <div>压缩包保存位置：<br /><input type="text" name="tarfile" id="tarfile" value="/bakup/备份_<%=Easp.Date.Format(Now, "ymmddhhiiss")%>.tar" /></div>
        <div><button id="pack">开始压缩打包到服务器指定位置</button> 或 <button id="down">直接通过浏览器下载压缩包</button></div>
        <div>
          
        </div>
      </div>
    </form>
  </fieldset>
  <fieldset>
    <legend>EasyASP 解压.tar格式演示</legend>
    <form method="post" id="form2">
      <div class="file-list" style="height:170px">
        <select id="files_bak" name="files_bak" size="20" style="height:170px"></select>
      </div>
      <div class="form1">
        <div>请选择要解压的.tar文件，用EasyASP压缩的或其它压缩软件(如好压)压缩的均可</div>
        <div>你选择的文件是：<br /><input type="text" name="untarfile" id="untarfile" value="" /></div>
        <div>解压后保存位置：<br /><input type="text" name="savepath" id="savepath" value="/bakup/" /></div>
        <div><input type="checkbox" name="hasself" id="hasself" checked="checked" value="1" style="width:20px;> <label for="hasself"">解压到独立的文件夹内</label></div>
        <div><button id="unpack">开始解压文件到服务器指定位置</button></div>
        <div>
          
        </div>
      </div>
    </form>
  </fieldset>
</body>
<script src="http://libs.baidu.com/jquery/1.10.2/jquery.min.js"></script>
<script type="text/javascript">
<!--
//显示左侧目录
loadtree('/', 0);
loadtree('/bakup/', 1);
//提交表单
$("#pack").on('click',function(){
  $('#form1').attr('action', 'tar.asp?act=pack');
  $("#form1").submit(form1submit);
});
$("#down").on('click',function(){
  $('#form1').attr('action', 'tar.asp?act=down');
  $("#form1").submit(form1submit);
});
$("#unpack").on('click',function(){
  $('#form2').attr('action', 'tar.asp?act=unpack');
  $("#form2").submit(form2submit);
});
//单击解压目录
$("#files_bak").on('click', function(){
  var v = $(this).val();
  if(v.substring(v.lastIndexOf('.'))=='.tar')
    $('#untarfile').val(v);
  else
    $('#untarfile').val('');
});
//验证选择
function form1submit(){
  if (!$("#files").val()){
    alert("必须选择目录或文件");
    return false;
  } else {
    return true;
  }
}
function form2submit(){
  if (!$("#untarfile").val()){
    alert("必须选择一个.tar文件");
    return false;
  } else {
    return true;
  }
}
//获取目录列表
function loadtree(folder, sel){
  $.post('tar.asp', {
      act : 'sitetree',
      root : escape(folder)
    }, function(data){
      if (folder!='/'){
        folder = folder.substring(0, folder.lastIndexOf('/'));
        folder = folder.substring(0, folder.lastIndexOf('/')+1);
      }else
        folder = null;
      showFiles(sel, data, folder);
    }, 
  'json');
}
//更新目录list数据
function showFiles(sel, list, fo){
  var s = sel==0 ? $("#files") : $("#files_bak");
  s.children().remove();
  if (fo){
    s.append('<option value="|'+fo+'" ondblclick="loadtree(this.value.substring(1), '+sel+')">..</option>');
  }
  for (var i in list){
    var opt = '<option value="'+list[i].name+'"';
    if (list[i].type=='file'){
      opt += ' style="color:#666"';
    } else {
      opt += ' style="background-color:#ff9"';
      opt += ' ondblclick="loadtree(this.value, '+sel+')"';
    }
    opt += '>'+list[i].name+'</option>';
    s.append(opt);
  }
}
//-->
</script>
</html>
<%
%>