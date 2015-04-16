<!--#include file="../../easyasp/easp.asp" --><%
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
str = str & "		<last><![CDATA[是不是<font color=""red"">这样</font>的噢，我也不知道哈。<img src=""http://img.t.sinajs.cn/t4/appstyle/expression/ext/normal/af/cry.gif"" />]]></last></site>" & vbCrLf
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
'Easp.PrintlnHtml Easp.Xml("name").Length
'Easp.Println "--------"
''选择所有包含属性alias的标签为name的节点
'Easp.PrintlnHtml Easp.Xml("name[alias]").Length
'Easp.Println "--------"
''选择所有属性for等于me，nick属性不等于email的标签为account的节点，并输出其Xml代码
'Easp.PrintlnHtml Easp.Xml("account[for='me'][nick!='email']").Xml
'Easp.Println "--------"
''选择site节点的子节点中标签为name的节点
'Easp.PrintlnHtml Easp.Xml("site>name").Xml
'Easp.Println "--------"
''选择account节点的后代节点中标签为name的节点
'Easp.PrintlnHtml Easp.Xml("account name").Xml
'Easp.Println "--------"
''选择所有的url和last节点
'Easp.PrintlnHtml Easp.Xml("url,last").Xml
Easp.Println "--------"
Easp.PrintlnHtml Easp.Xml("url")(2).Xml
Easp.Xml("url")(2).Text = "<test>sss</test>"
Easp.PrintlnHtml Easp.Xml("url")(2).Xml

'Easp.Xml.XSLT = "xsl/microblog.xsl"
'Easp.PrintlnHtml Easp.Xml.Dom.Xml

'Easp.Println Easp.Xml.SaveAs("news.xml>gbk")
'Easp.Println Easp.Xml.SaveAs("microblog.xml>utf-8")

'Set n = Easp.Xml("title")
'For i = 0 To n.Length-1
'	Easp.Println n(i).Value
'Next
'Set n = Nothing

'Easp.Println Easp.Xml("last")(2).Value
Set n = Easp.Xml("last")
For i = 0 To n.Length-1
	Easp.Println n(i).Type
	Easp.Println n(i).Value
Next
'Easp.Println n.Text
'Easp.Println n(1).Root.Type
'Easp.Println n(2).Parent.Name
'Easp.Println n(0).Clone(1).Text
'Set n = Nothing
'Easp.Xml("name")(0).RemoveAttr("alias")
'Easp.PrintlnHtml Easp.Xml("name")(0).Xml
'Easp.Xml("site")(1).Clear
'Easp.PrintlnHtml Easp.Xml("site")(1).Xml

'Easp.PrintlnHtml TypeName(Easp.Xml("site")(0).Parent.Parent.Dom)
'Easp.Xml("url").Remove
'Easp.Xml("name").Attr("alias") = Null
'Easp.Xml("microblog").Remove
'Easp.Println Easp.Xml.Sel("//site").Length
'Easp.Println Easp.Xml.Select("//site").Length
'Easp.Println Easp.Xml("site").Length
'Easp.Println Easp.Xml("site").Type
'Easp.Xml("url")(2).Value = "http://sss.com"
'Easp.Println TypeName(n)
'替换节点
'Set n = Easp.Xml("name")(1).ReplaceWith(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("name").ReplaceWith(Easp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Easp.Xml("name")(1).ReplaceWith(Easp.Xml("url")(2))
'Easp.PrintlnHtml n.Xml
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


'Easp.PrintlnHtml n.Xml
'Easp.PrintlnHtml Easp.Xml.Dom.Xml

'Easp.PrintlnHtml Easp.Xml("name").Length
'Easp.PrintlnHtml Easp.Xml("site name").Length
'Easp.PrintlnHtml Easp.Xml("site>name").Length
'Easp.PrintlnHtml Easp.Xml("name[alias='Tencent'],url").Length
'Easp.PrintlnHtml Easp.Xml("name[alias='Tencent'],url").Text
'Easp.PrintlnHtml Easp.Xml.Select("//account[@nick='user' and position()<2]").Length
'Easp.PrintlnHtml Easp.Xml.Select("//account[@nick='user' and position()<2]").Xml
'Easp.PrintlnHtml Easp.Xml("account[nick='user'][for!='me'],account[nick!='user']").Xml

'Easp.PrintlnHtml Easp.Xml("site")(1).Find("account").Root.TypeString
'Easp.PrintlnHtml Easp.Xml.Root.TypeString

'Set n = Nothing
Easp.Println ""
Easp.Println "------------------------------------"
Easp.Print "页面执行时间： " & Easp.GetScriptTime & " 秒"
%>