<!--#include file="../../easyasp/easp.asp" --><%

If Easp.Has(Easp.Get("action")) Then
  Dim act
  '验证url参数，必须等于 save
  act = Easp.GetVal("action").Name("action").Same("save").Alert
  Easp.Var("username") =  Easp.PostVal("username").Name("用户名").Required.Test("username").Alert
  Easp.Var("email") =  Easp.PostVal("email").Name("Email").Test("email").Alert
  '验证两次输入的密码一致
  Easp.VarVal("post.password").Name("密码").Required.Trim().MinLength(6).SamePost("passwordrepeate").Alert
  Easp.Var("password") = Easp("md5")(Easp.Post("password"))
  '验证序列
  Easp.Var("type") = Easp.VarVal("post.type").Name("类型").Split(", ").IsNumber.Join("|").Alert()
  '验证验证码
  Session("verifycode") = "E92A"
  Easp.Println "Verify Code:" & Easp.PostVal("verify").SameSession("verifycode").Msg("wrong code").PrintEndJson()
  Easp.Println "UserName:" & Easp.Var("username")
  Easp.Println "Password:" & Easp.Var("password")
  Easp.Println "Type:" & Easp.Var("type")
End If
%>
<form action="?action=save&username=coldstone" method="post">
  username: <input type="text" size="60" name="username" value="" /><br />
  email: <input type="text" size="60" name="email" value="" /><br />
  password: <input type="password" size="60" name="password" value="" /><br />
  repeat: <input type="password" size="60" name="passwordrepeate" value="" /><br />
  verify: <input type="text" size="20" name="verify" value="" />E92A<br />
  <input type="checkbox" name="type" value="1" checked="checked" />type1
  <input type="checkbox" name="type" value="2" />type2
  <input type="checkbox" name="type" value="3" />type3
  <input type="checkbox" name="type" value="4" />type4<br />
  <button type="submit">Submit to "?action=save&username=coldstone"</button>
</form>