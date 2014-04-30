<!--#include file="../../easyasp/easp.asp" --><%

Session("StringTest") = "This is a test string for Easp.Str.ToString"

Easp.SetCookie "app_name", "easyAsp_Name", ""
Easp.SetCookie "site>my_name", "coldstone_easp", ""
Easp.SetCookie "site>mytype", "very&diaosi", ""

Easp.Println "Easp.Str.IsSame(""ABCD"", ""abcd"") : " & Easp.Str.IsSame("ABCD", "abcd")
Easp.Println "Easp.Str.IsEqual(""ABCD"", ""abcd"") : " & Easp.Str.IsEqual("ABCD", "abcd")
Easp.Println "Easp.Str.IsEqual(""ABCD"", ""<"", ""abcd"") : " & Easp.Str.Compare("ABCD", "<", "abcd")
Easp.Println ""
Easp.Println "Easp.Str.GetColonName(""username"") : " & Easp.Str.GetColonName("username")
Easp.Println "Easp.Str.GetColonName(""username:value"") : " & Easp.Str.GetColonName("username:value")
Easp.Println "Easp.Str.GetColonValue(""username"") : " & Easp.Str.GetColonValue("username")
Easp.Println "Easp.Str.GetColonValue(""username:value"") : " & Easp.Str.GetColonValue("username:value")
Easp.Println "Easp.Str.GetColonValue(""username:"") : " & Easp.Str.GetColonValue("username:")
Easp.Println ""
Easp.Println "Easp.Str.ToString(""testing"") : " & Easp.Str.ToString("testing")
Easp.Println "Easp.Str.ToString(Array(""yes"",""no"",""unknown"")) : " & Easp.Str.ToString(Array("yes","no","unknown"))
Easp.Println "Easp.Str.ToString(Array(12,34,111,98,0)) : " & Easp.Str.ToString(Array(12,34,111,98,0))
Easp.Println "Easp.Str.ToString(Array()) : " & Easp.Str.ToString(Array())
Easp.Println "Easp.Str.ToString(Empty) : " & Easp.Str.ToString(Empty)
Easp.Println "Easp.Str.ToString(Null) : " & Easp.Str.ToString(Null)
Easp.Println "Easp.Str.ToString(Nothing) : " & Easp.Str.ToString(Nothing)
Easp.Println "Easp.Str.ToString(Err) : " & Easp.Str.ToString(Err)
Easp.Println ""
Dim dic : Set dic = Server.CreateObject("Scripting.Dictionary")
Easp.SetDictionaryKey dic, "my-time", Now
Easp.SetDictionaryKey dic, "my", "Yes,it's me."
Easp.Println "Easp.Str.ToString(dic) : " & Easp.Str.ToString(dic)
Easp.Println ""
Easp.Var("EaspVar") = "Easyasp variable"
Easp.Var("TestArray") = Array(2323,490,108,"我是中文",992,83,920)
Easp.Println "Easp.Str.ToString(Easp.Var.GetObject) : " & Easp.Str.ToString(Easp.Var.GetObject)
'Easp.Console Easp.Var.GetObject
Easp.Println ""
Easp.Println "Easp.Str.ToString(12.1256) : " & Easp.Str.ToString(12.1256)
Easp.Println "Easp.Str.ToString(12.100) : " & Easp.Str.ToString(12.100)
Easp.Println "Easp.Str.ToString(12.005) : " & Easp.Str.ToString(12.005)
Easp.Println "Easp.Str.ToString(12.00) : " & Easp.Str.ToString(12.00)
Easp.Println ""
Easp.Println "Easp.Str.ToString(Session) : " & Easp.Str.ToString(Session)
Easp.Println "Easp.Str.ToString(Request.Cookies) : " & Easp.Str.ToString(Request.Cookies)
Easp.Println "Easp.Str.ToString(Request.QueryString) : " & Easp.Str.ToString(Request.QueryString)
Easp.Println "Easp.Str.ToString(Request.Form) : " & Easp.Str.ToString(Request.Form)
Easp.Println ""
Easp.Println "Easp.Str.Cut(""This我"",3) : " & Easp.Str.Cut("This我",3)
Easp.Println "Easp.Str.Cut(""我是一个人"",4) : " & Easp.Str.Cut("我是一个人",4)
Easp.Println ""
Easp.Println "Easp.Str.RepPart(""photo-3.html"", ""^(\w+)-(\d+)\.html$"", ""$2"", ""4"") : " & Easp.Str.ReplacePart("photo-3.html", "^(\w+)-(\d+)\.html$", "$2", "4")
Easp.Println ""
Easp.Println "Easp.Str.RandomNumber(1000,9999) : " & Easp.Str.RandomNumber(1000,9999)
Easp.Println "Easp.Str.RandomStr(10) : " & Easp.Str.RandomStr(10)
Easp.Println "Easp.Str.RandomStr(""12:0123456789abcdefghijklmnopqrstuvwxyz~!@#$%^&*_-+="") : " & Easp.Str.RandomStr("12:0123456789abcdefghijklmnopqrstuvwxyz~!@#$%^&*_-+=")
Easp.Println "Easp.Str.RandomStr(""10000-99999"") : " & Easp.Str.RandomStr("10000-99999")
Dim color : color = Easp.Str.RandomStr("#<3>:0123456789ABCDEF")
Easp.Println "Easp.Str.RandomStr(""Random Color \: #<3>:0123456789ABCDEF"") : <span style=""background-color:" & color & """>Random Color : " & color & "</span>"
Easp.Println "Easp.Str.RandomStr(""{<8>-<4>-<4>-<4>-<12>}:0123456789ABCDEF"") : " & Easp.Str.RandomStr("{<8>-<4>-<4>-<4>-<12>}:0123456789ABCDEF")
Easp.Println "Easp.Str.RandomStr(""CN-\<86\>-<6>-<10000-99999>"") : " & Easp.Str.RandomStr("CN-\<86\>-<6>-<10000-99999>")
Easp.Println ""
Easp.Println "Easp.Str.ToNumber(number, decimalType) 方法："
Easp.Println "如果第二个参数为N，则保留N位小数，小数位数不足的补0"
Easp.Println "Easp.Str.ToNumber(0.345678, 3) : " & Easp.Str.ToNumber(0.345678, 3)
Easp.Println "Easp.Str.ToNumber(0.34, 3) : " & Easp.Str.ToNumber(0.34, 3)
Easp.Println "Easp.Str.ToNumber(0, 3) : " & Easp.Str.ToNumber(0, 3)
Easp.Println "如果第二个参数为0，则保留所有小数位数"
Easp.Println "Easp.Str.ToNumber(0.345678, 0) : " & Easp.Str.ToNumber(0.345678, 0)
Easp.Println "Easp.Str.ToNumber(0.34, 0) : " & Easp.Str.ToNumber(0.34, 0)
Easp.Println "Easp.Str.ToNumber(0, 0) : " & Easp.Str.ToNumber(0, 0)
Easp.Println "如果第二个参数为-N，则保留N位小数，但小数位数不足的不补0"
Easp.Println "Easp.Str.ToNumber(0.345678, -3) : " & Easp.Str.ToNumber(0.345678, -3)
Easp.Println "Easp.Str.ToNumber(0.34, -3) : " & Easp.Str.ToNumber(0.34, -3)
Easp.Println "Easp.Str.ToNumber(0, -3) : " & Easp.Str.ToNumber(0, -3)
Easp.Println ""
Easp.Println "Easp.Str.ToPrice(12.3456) : " & Easp.Str.ToPrice(12.3456)
Easp.Println "Easp.Str.ToPercent(0.3456) : " & Easp.Str.ToPercent(0.3456)
Easp.Println ""
Easp.Println "Easp.Str.Half2Full(""半角To全角"") : " & Easp.Str.Half2Full("半角To全角")
Easp.Println "Easp.Str.Full2Half(""全角Ｔｏ半角"") : " & Easp.Str.Full2Half("全角Ｔｏ半角")
%>