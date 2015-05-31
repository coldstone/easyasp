<!--#include file="../../easyasp/easp.asp" --><%
'测试解析JSON
Dim s_json, obj
s_json = Easp.Fso.Read("sample.json")
Set obj = Easp.Decode(s_json)
Dim Devices,i
'如果是数组需要用.GetArray方法取出后方可循环
Devices = obj("Circuit[0].Devices").GetArray
'Devices = Devices
For i = 0 To Ubound(Devices)
  Easp.Print "Name:"
  Easp.Println Devices(i)("Name")
  Easp.Print "dSID:"
  Easp.Println Devices(i)("dSID")
  Easp.Print "ZoneID:"
  Easp.Println Devices(i)("ZoneID")
  Easp.Println "======="
Next
Set obj = Nothing
Easp.Println "=============================="
s_json = Easp.Fso.Read("samplewithcomment.json")
Set obj = Easp.Decode(s_json)
'有两种方式访问解析后的Json对象或数组
Easp.Println obj(0)("alert")("message")(1)("set")("name")
Easp.Println obj(0)("alert.message[2].switch.case.input.title")
Easp.Println "=============================="
Easp.PrintlnString Easp.Encode(obj)
%>