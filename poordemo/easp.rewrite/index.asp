<!--#include file="../../easyasp/easp.asp" --><%

'转到带地址栏的链接
Easp.Print "<a href=""?type=Easp.coldstone&id=1983-09-23&page=3&lang=%E4%B8%AD%E6%96%87"">set address querystring</a>"
Easp.Print "&nbsp;&nbsp;&nbsp;"
Easp.Print "<a href=""./index.asp?photo-203HTKJI9B-6.html"">set address rewrite</a>"
Easp.Print "&nbsp;&nbsp;&nbsp;"
Easp.Println "<a href=""./?photo-203HTKJI9B-6.html"">set address rewrite without 'index.asp'</a>"

Easp.Println "Easp.DefaultPageName : " & Easp.DefaultPageName
Easp.Println "[All] Easp.GetUrl("""") : " & Easp.GetUrl("")
Easp.Println "[Url] Easp.GetUrl(1) : " & Easp.GetUrl(1)
Easp.Println "[Url] Easp.GetUrl(0) : " & Easp.GetUrl(0)
Easp.Println "[Host] Easp.GetUrl(-1) : " & Easp.GetUrl(-1)
Easp.Println "[Dir] Easp.GetUrl(-2) : " & Easp.GetUrl(-2)
Easp.Println "[File] Easp.GetUrl(-3) : " & Easp.GetUrl(-3)
Easp.Println "[White] Easp.GetUrl(""type,id"") : " & Easp.GetUrl("type,id")
Easp.Println "[Black] Easp.GetUrl(""-type,-id"") : " & Easp.GetUrl("-type,-id")
Easp.Println "[Remove all param] Easp.GetUrl(""-:all"") : " & Easp.GetUrl("-:all")
Easp.Println "[New param] Easp.GetUrlWith(""-page,-lang"", ""page=4&lang=english"") : " & Easp.GetUrlWith("-page,-lang", "page=4&lang=english")
Easp.Println "[New page & param] Easp.GetUrlWith(""./newpage.asp?-type,-id,-lang"", ""lang=english"") : " & Easp.GetUrlWith("./newpage.asp?-type,-id,-lang", "lang=english")
Easp.Println ""
'设置伪静态规则
'Easp.RewriteRule "/testcase/rewrite/\?(\w+)-(\w+)-(\d+).html", "/testcase/rewrite/?type=$1&id=$2&page=$3"
'另一种方式设置伪静态规则
'Easp.Rewrite "/testcase/rewrite/index.asp", "(\w+)-(\w+)-(\d+).html", "type=$1&id=$2&page=$3"
'设置本页面伪静态规则
Easp.Println "Easp.Rewrite """", ""(\w+)-(\w+)-(\d+).html"", ""type=$1&id=$2&page=$3"""
Easp.Rewrite "", "(\w+)-(\w+)-(\d+).html", "type=$1&id=$2&page=$3"

Easp.Println "当前页是否符合伪静态规则：" & Easp.IsRewrite()

Easp.Println "输出参数值："
Easp.Println "Easp.Get(""type"") : " & Easp.Get("type")
Easp.Println "Easp.Get(""id"") : " & Easp.Get("id")
Easp.Println "Easp.Get(""page"") : " & Easp.Get("page")
Easp.Println "替换URL参数值："
Easp.Println "Easp.ReplaceUrl(""page"", 2) : " & Easp.ReplaceUrl("page", 2)
Easp.Println "Easp.ReplaceUrl(""class"", 2) : " & Easp.ReplaceUrl("class", 2)

Easp.Println "Easp.Str.ToString(Easp.Var.GetObject) : " & Easp.Str.ToString(Easp.Var.GetObject)

%>