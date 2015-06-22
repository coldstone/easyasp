<!--#include file="../../easyasp/easp.asp" --><%
Dim txt,i,j,html,tpl
txt = "这篇文档是Easp的模板类的测试文档和示例文件"

Set tpl = Easp.Tpl.New
'允许在模板文件中使用ASP代码
tpl.AspEnable = True

'模板文件所在文件夹，支持绝对路径和相对路径
tpl.FilePath = "html/"
'模板文件中可以用{#include}标签包含无限层次的子模板，也都支持相对路径和绝对路径，请参考html文件夹内的模板文件

'如何处理未替换的标签,"keep"-保留，"remove"-移除，"comment"-转成注释
'tpl.TagUnknown = "comment"

'模板标签的样式，默认为"{*}"，*号为标签名
'tpl.TagMask = "{$*$}"

'加载模板
tpl.Load "tpl.html"

'也可以用下面这种方式加载模板
'tpl.File = "tpl.html"

'开始解析标签，MakeTag可以快速生成html标签
tpl "author", tpl.MakeTag("author","Coldstone, TainRay")
tpl "keywords", tpl.MakeTag("keywords", "EasyAsp, Easp, Version 2.2")
tpl "description", tpl.MakeTag("description","This is a EasyAsp TPL Sample.")

'将标签替换为副模板
tpl.TagFile "style", "inc/style.html"

tpl "jsfile", tpl.MakeTag("js","html/inc.js")
tpl "cssfile", tpl.MakeTag("css","html/style.css")
tpl "title", "EasyAsp 模板类测试页"
tpl "subtitle", txt
tpl "color", "#F60"

If Hour(Now)>=10 Then
	'追加标签内容：
	tpl.Append "subtitle", " <small>[10点之后显示]</small>"
End If

'开始循环
For i = 1 to 3
	tpl "A.title", "A标题" & i
	tpl "A.addtime", Now + i
	'更新本次循环数据，每次循环后必须调用此方法
	tpl.Update "A"
Next

'嵌套循环演示，嵌套可以无限层的，这是父循环
For i = 1 to Easp.Str.RandomNumber(3,6)
	tpl "B.title", "B标题" & i & " | "
	'Demo中的 B. 这个前缀不是必须的，只是为了代码方便阅读
	tpl "id", i+10
	tpl "addtime", Now + i
	'这是子循环
	For j = 20 to Easp.Str.RandomNumber(22,25)
		'替换标签
		tpl "page.list", " "&i&">"&j
		'更新本次循环数据
		tpl.Update "page"
	Next
	'更新本次循环数据
	tpl.Update "B"
Next

'将替换完毕的html输出至浏览器
tpl.Show

'或者，也可以生成静态页
'得到替换完毕的html代码
html = tpl.GetHtml
'生成静态页
'Call Easp.Fso.CreateFile("demo.tpl.html",html)
%>