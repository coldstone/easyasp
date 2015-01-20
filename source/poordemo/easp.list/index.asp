<!--#include file="../../easyasp/easp.asp" --><%
Easp.BasePath = "/source/easyasp"
'先构造一个随机数组，给下面的第1种方法用：
Dim arrayA(19),i
For i = 0 To 19
	arrayA(i) = Easp.Str.RandomStr(Easp.Str.RandomNumber(0,6)&":abcdeABCDE1234567890")
Next
Dim list, Alist, arr
'创建一个List对象
Set list = Easp.List.New
'忽略大小写(去重复项、搜索、取索引值、排序、比较时)
'list.IgnoreCase = False
'-----------------------------------------------
'把数组存入List对象管理(可以用2种方法接受共4种形式的数据)
'-------------
'第1种：简单数组
'list.Data = arrayA
'第2种：用空格隔开的字符串，每个字符串会解析为数组的一个元素
'list.Data = "aa a ee ddd A AA aa Ab ab a bb b  bb c ccc  ddd b d"
'-------------
'第3种：带下标的数组，如果数组元素中包含 : 号，则会解析为Hash表值对， : 号前的字符串为Hash的下标
'list.Hash = Array("test:34", "name:coldstone", 344.89, "birth:81/01/01", "Others", "btime:81/01/32", "addtime:"&True)
'第4种：用空格隔开的字符串，字符串中包含 : 号，也会把带 : 号的字符串解析为Hash表值对
list.Hash = "aa a ee se:ddd A AA aa my:Ab ab a bb b  bb c la:ccc  ddd b d"
'-----------------------------------------------
Easp.Println "初始数组为：" & list.ToString

list("one") = "OneNumber"
list("two") = "2222"
list("six") = "SSSSix"
'删除第一个元素
list.Shift
'添加一个元素到开头
list.UnShift "first"
'在结尾添加一个元素
list.Push "wobu"
list.Push "zhidao"
list.Push -349.89
list.Push 80
list.Push "ssssix"
'插入元素
list.Insert 22, Array("seven","eight","nine")
Easp.Println "添加一些元素(包括Hash表值)后为：" & list.ToString
Easp.Println "所有的Hash表值都可以同时用Hash名称下标和数字下标取值(就像记录集的字段那样)，如：list(""one"") = " & list("one") & " ，list(20) = " & list(20)
'删除最后一个元素
list.Pop
'删除指定元素
list.Delete 4
list.Delete "two"
Easp.Println "删除一些元素后为：" & list.ToString
Easp.Println "现在数组的长度是：" & list.Size
Easp.Println "数组的有效值个数（非空值）是：" & list.Count
'去除重复元素
list.Uniq
Easp.Println "去除重复元素后为：" & list.ToString
list.Compact
Easp.Println "去除空元素后为：" & list.ToString
Easp.Println "数组的最大值是：" & list.Max
Easp.Println "数组的最小值是：" & list.Min
Easp.Println "数组的第一个元素是：" & list.First
Easp.Println "数组的最后一个元素是：" & list.Last
Easp.Println "是否包含下标为 ""six"" 的元素：" & list.HasIndex("six") & "， 它的数字下标是： " & list.Index("six")
Easp.Println "把包含的Hash表对值转化为url参数字符串为：" & list.Serialize
Easp.Println "=========="
list.Sort
Easp.Println "排序后为：" & list.ToString
list.Reverse
Easp.Println "倒序后为：" & list.ToString
list.Rand
Easp.Println "打乱顺序后为：" & list.ToString
Easp.Println "是否包含字符串 ""bb"" ：" & list.Has("bb") & "，在数组中第1次出现的下标是：" & list.IndexOf("bb")
Easp.Println "=========="
Easp.Println ""
Easp.Println "所有可以操作数组的方法后加 _ 则是返回一个新的数组对象，不会改变原数组的数据："
Easp.Println "--------"
Easp.Println "数组中包含字符串 ""a"" 的元素(不影响原数组)：" & list.Search_("a").ToString
Easp.Println "数组中不包含字符串 ""a"" 的元素(不影响原数组)：" & list.SearchNot_("a").ToString
Easp.Println "执行迭代处理(不影响原数组)：" & list.Map_("testmy").ToString
'按条件选择
Easp.Println "第一个是数字的值是：" & list.Find("isNumeric(%i)")
Easp.Println "选择所有非数字的值(不影响原数组)：" & list.Select_("Not isNumeric(%i)").ToString
'按正则表达式选择
Easp.Println "选择所有以数字开头的值(不影响原数组)：" & list.Grep_("^\d.+").ToString
'迭代：依次用函数处理每个值并返回
Easp.Println "执行迭代处理后排序(不影响原数组)：" & list.SortBy_("testmy").ToString
Easp.Println ""
'迭代处理：依次用数组作为参数值执行函数
Easp.Println "==以下是迭代执行函数：=="
list.Each("toUp")
Easp.Println "==迭代执行函数结束=="
Easp.Println ""

'迭代用的函数
Function testmy(ByVal s)
	testmy = "U : " & UCase(s)
End Function

'迭代用的函数
Sub toUp(ByVal s)
	Easp.Println s & " ==&gt; " & UCase(s)
End Sub

'取得数组的其中一部分（按下标）
'可取多个元素，用逗号隔开，可以用 - 表示范围（如2-5表示第2到第5下标，\s表示开头，\e表示结尾, list.Delete也可以用这样的下标删除）
list.Slice "1,3,6-\e"
Easp.Println "取下标为 ""1,3,6-结束"" 的结果用 | 连起来是：" & list.Join(" | ")

'数组重复
Easp.Println "数组重复2遍后的结果(不影响原数组)：" & list.Times_(2).ToString

Set Alist = Easp.List.New
'Alist.Hash = "aaa:ssssix b:wefewr c:sfwef one:weioid six:yesterday ee"
Alist.Data = Array("ssssix","OneNumber","zhidao",234.234,35235,3534.345,78)
'附加数组(参数可以是Array数组，也可以是List对象)
Easp.Println "附加一个数组后的结果(不影响原数组)：" & list.Splice_(Alist).ToString
'合并数组
Easp.Println "合并数组后的结果(不影响原数组)：" & list.Merge_(Alist).ToString
'数组交集
Easp.Println "取数组交集后的结果(不影响原数组)：" & list.Inter_(Alist).ToString
'数组差集
Easp.Println "取数组差集后的结果(不影响原数组)：" & list.Diff_(Alist).ToString
'比较数组的大小
Easp.Println "数组比较的结果(1：大于，-1小于，0等于)：" & Alist.Eq(Array("ssssix","OneNumber","zhidao1",234.234,35235,3534.345))
'数组是否包含另一数组
Easp.Println "是否包含另一数组的结果：" & Alist.Son(Alist.Pop_)
Set Alist = Nothing

Easp.Println "=========="
Easp.Println "---遍历现在的List---"
For i = 0 To list.End
	Easp.Println "list("&i&") 的值是：" & list(i)
Next
Easp.Println "=========="
Easp.Println "---遍历现在的List中的散列对值---"
Dim Maps,key,x,y
Set Maps = list.Maps
For Each key In Maps
	If Not isNumeric(key) Then
		Easp.Println "list(""" & key & """) = list(" & Maps(key) &  ") = " & list(key)
	End If
Next
Set Maps = Nothing

Easp.Println "=========="

Easp.Println "---取出为普通数组(如果是Hash表值就把Hash名称转换为前缀带:)后再遍历---"
arr = list.Hash
For i = 0 To Ubound(arr)
	Easp.Println "arr("&i&") 的值是：" & arr(i)
Next

Easp.Println "=========="

Easp.Println "---取出为普通数组后再遍历---"
arr = list.Data
For i = 0 To Ubound(arr)
	Easp.Println "arr("&i&") 的值是：" & arr(i)
Next

Easp.Println "------------------------------------"
Easp.Print "页面执行时间： " & Easp.GetScriptTime & " 秒"
Set list = Nothing
Set Easp = Nothing
%>