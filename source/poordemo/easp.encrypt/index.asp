<!--#include file="../../easyasp/easp.asp" --><%
Dim i_type, s_source, s_target
If Easp.Get("action") = "en" Then
  i_type = Easp.Post("type")
  Easp.Encrypt.Key = Easp.Post("key")
  If i_type = 0 Then
    s_source = Easp.Post("source")
    s_target = Easp.Encrypt(s_source)
    Easp.PrintEnd s_target
  ElseIf i_type = 1 Then
    s_target = Easp.Post("target")
    s_source = Easp.Encrypt.Decrypt(s_target)
    Easp.PrintEnd s_source
  End If
End If
%>
<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8" />
	<title>Easp.Encrypt</title>
	<style>
  	div{text-align:center;}
  	.container{width:90%;margin:0 auto;}
  	.container textarea {width:100%;height:320px;border:1px solid #222;word-break:break-all;}
  	.btns{padding:10px;}
	</style>
</head>
<body>
	<div class="container">
  	<div>
    	<textarea name="source" id="source">EasyASP 加密解密类 Easp.Encrypt 采用了一种高效率的对称加密算法，用于普通的加解密场景应该是足够了，密钥设置得越长安全性越高，优点是速度非常快。
使用方法
&lt;%
'设置密钥
Easp.Encrypt.Key = "encrypt_key"
'加密
Easp.Encrypt(string)
'解密
Easp.Encrypt.Decrypt(string)
'用指定密钥加密
Easp.Encrypt.EncryptBy(string, "encrypt_key")
'用指定密钥解密
Easp.Encrypt.DecryptBy(string, "encrypt_key")
%&gt;
Unicode字符示例：
ワールドカップ ブラジル大会 : ブラジル監督 ネイマール帯同を望む
메탈리카, 글래스톤베리 달궈 : 생생 팝 단신 '지구촌 팝뉴스'
قصة بالصور: يوم من أيام رمضان في أسرة من شينجيانغ
</textarea>
    </div>
  	<div class="btns">密钥：<input type="text" name="key" id="key" value="encrypt_key"> <input id="btn_en" type="button" value="加密(Encrypt) ↓"> <input id="btn_de" type="button" value="解密(Decrypt) ↑"></div>
  	<div><textarea name="target" id="target"></textarea></div>
	</div>
</body>
<script src="//lib.sinaapp.com/js/jquery/1.10.2/jquery-1.10.2.min.js"></script>
<script type="text/javascript">
<!--
  $(function(){
    $('#btn_en').on('click',function(){encrypt(true)});
    $('#btn_de').on('click',function(){encrypt(false)});
  });
  function encrypt(type){
    type = type ? 0 : 1;
    $('#' + ['target','source'][type]).val('loading...');
    $.post(
      'index.asp?action=en',
      {type:type, key:$('#key').val(), source:$('#source').val(), target:$('#target').val()},
      function(data){
        if(type==0){
          $('#target').val(data);
        } else if(type==1) {
          $('#source').val(data);
        }
      }
    );
  }
//-->
</script>
</html>
