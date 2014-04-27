<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Debug = True
Easp.Use "Xml"
Dim str,n,i
str = 			"<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
str = str & "<microblog>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Tencent"">腾讯微博</name>" & vbCrLf
str = str & "		<url>http://t.qq.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""me""><name>@lengshi</name><nick>Ray</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[今天我们这里下<em>大雨</em>啦！]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Sina"">新浪微博</name>" & vbCrLf
str = str & "		<url>http://t.sina.com.cn</url>" & vbCrLf
str = str & "		<account nick=""email"" for=""me""><name>@tainray</name><nick>tainray@sina.com</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[是不是<font color=""red"">这样</font>的噢，我也不知道哈。<img src=""http://bbs.lengshi.com/max-assets/icon-emoticon/12.gif"" />]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Twitter"">推特</name>" & vbCrLf
str = str & "		<url>http://twitter.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""notme""><name haha=""1"">@ccav</name><nick>CCAV</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[I don't need this feature <strong>(>_<)</strong> any more.]]></last></site>" & vbCrLf
str = str & "</microblog>"

'载入Xml数据
'Easp.Xml.Load "http://easp.lengshi.cn/data/xml/microblog_catalog.xml"
Easp.Xml.Load str
''选择所有标签为name的节点，并输出找到的节点个数
'Easp.WNH Easp.Xml("name").Length
'Easp.WN "--------"
''选择所有包含属性alias的标签为name的节点
'Easp.WNH Easp.Xml("name[alias]").Length
'Easp.WN "--------"
''选择所有属性for等于me，nick属性不等于email的标签为account的节点，并输出其Xml代码
'Easp.WNH Easp.Xml("account[for='me'][nick!='email']").Xml
'Easp.WN "--------"
''选择site节点的子节点中标签为name的节点
'Easp.WNH Easp.Xml("site>name").Xml
'Easp.WN "--------"
''选择account节点的后代节点中标签为name的节点
'Easp.WNH Easp.Xml("account name").Xml
'Easp.WN "--------"
''选择所有的url和last节点
'Easp.WNH Easp.Xml("url,last").Xml
Easp.WN "--------"
Easp.WNH Easp.Xml("url")(2).Xml
Easp.Xml("url")(2).Text = "<test>sss</test>"
Easp.WNH Easp.Xml("url")(2).Xml

'Easp.Xml.XSLT = "xsl/microblog.xsl"
'Easp.WNH Easp.Xml.Dom.Xml

'Easp.WN Easp.Xml.SaveAs("news.xml>gbk")
'Easp.WN Easp.Xml.SaveAs("microblog.xml>utf-8")

'Set n = Easp.Xml("title")
'For i = 0 To n.Length-1
'	Easp.WN n(i).Value
'Next
'Set n = Nothing

'Easp.WN Easp.Xml("last")(2).Value
Set n = Easp.Xml("last")
For i = 0 To n.Length-1
	Easp.WN n(i).Type
	Easp.WN n(i).Value
Next
'Easp.WN n.Text
'Easp.WN n(1).Root.Type
'Easp.WN n(2).Parent.Name
'Easp.WN n(0).Clone(1).Text
'Set n = Nothing
'Easp.Xml("name")(0).RemoveAttr("alias")
'Easp.WNH Easp.Xml("name")(0).Xml
'Easp.Xml("site")(1).Clear
'Easp.WNH Easp.Xml("site")(1).Xml

'Easp.WNH TypeName(Easp.Xml("site")(0).Parent.Parent.Dom)
'Easp.Xml("url").Remove
'Easp.Xml("name").Attr("alias") = Null
'Easp.Xml("microblog").Remove
'Easp.WN Easp.Xml.Sel("//site").Length
'Easp.WN Easp.Xml.Select("//site").Length
'Easp.WN Easp.Xml("site").Length
'Easp.WN Easp.Xml("site").Type
'Easp.Xml("url")(2).Value = "http://sss.com"
'Easp.WN TypeName(n)
'替换节点
'Set n = Easp.Xml("name")(1).ReplaceWith(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("name").ReplaceWith(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("name")(1).ReplaceWith(Easp.Xml("url")(2))
'Easp.WNH n.Xml
'清空
'Easp.Xml("url").Empty
'Easp.Xml("name").Clear
'从前面加入节点
'Set n = Easp.Xml("account")(1).Before(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("account")(1).Before(Easp.Xml("url")(2))
'Set n = Easp.Xml("account").Before(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("account").Before(Easp.Xml("url")(2))
'从后面加入节点
'Set n = Easp.Xml("account")(2).After(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("last")(1).After(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("account")(1).After(Easp.Xml("url")(2))
'Set n = Easp.Xml("account").After(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("account").After(Easp.Xml("url")(2))


'Easp.WNH n.Xml
'Easp.WNH Easp.Xml.Dom.Xml

'Easp.WNH Easp.Xml("name").Length
'Easp.WNH Easp.Xml("site name").Length
'Easp.WNH Easp.Xml("site>name").Length
'Easp.WNH Easp.Xml("name[alias='Tencent'],url").Length
'Easp.WNH Easp.Xml("name[alias='Tencent'],url").Text
'Easp.WNH Easp.Xml.Select("//account[@nick='user' and position()<2]").Length
'Easp.WNH Easp.Xml.Select("//account[@nick='user' and position()<2]").Xml
'Easp.WNH Easp.Xml("account[nick='user'][for!='me'],account[nick!='user']").Xml

'Easp.WNH Easp.Xml("site")(1).Find("account").Root.TypeString
'Easp.WNH Easp.Xml.Root.TypeString

'Set n = Nothing
Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>