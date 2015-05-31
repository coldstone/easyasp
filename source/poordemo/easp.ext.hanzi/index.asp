<!--#include file="../../easyasp/easp.asp" --><!--#include file="../../easyasp/plugin/easp.hanzi.asp" --><%
Easp.BasePath = "/source/easyasp"
'测试汉字转拼音插件
Dim Hanzi, cn
cn = "二哥，传说嘉陵江边的重庆人都是重口味，跟《麻辣烫》中一样一样的。"
'Set Hanzi = Easp.Ext("Hanzi")
Set Hanzi = New EasyASP_Hanzi
Easp.Println "TEXT : " & cn
Easp.Println "TitleCase : " & Hanzi.TitleCase
''设置为首字母不大写
'Hanzi.FirstLetterUpcase = False
Easp.Println "GetPinYin : " & Hanzi.GetPinYin(cn)
Easp.Println "GetPY : " & Hanzi.GetPY(cn)
Easp.Println "GetPinYinRead : " & Hanzi.GetPinYinRead(cn)
Easp.Println "GetPinyin1234 : " & Hanzi.GetPinyin1234(cn)
'GetPinYinWith("中文字符串", 拼音韵母转为字母, 拼音后标识声调, 拼音间加空格, 仅取首字母, 首字母大写)
Easp.Println "GetPinYinWith : " & Hanzi.GetPinYinWith(cn, True, False, True, False, True)
Easp.Println "GetEnglish : " & Hanzi.GetEnglish(cn)
Easp.Println "GetEnglishDash : " & Hanzi.GetEnglishDash(cn)
Easp.Println "GetKeyWord : " & Hanzi.GetKeyWord(cn)
Easp.Println "GetKeyWordArray : " & Easp.Encode(Hanzi.GetKeyWordArray(cn))
Easp.Println "============================"
Easp.Println Easp.GetScriptTime & "s"

%>