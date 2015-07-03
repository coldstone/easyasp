<%
'######################################################################
'## easp.list.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP List(Array) Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2015-02-03 1:18:39
'## Description :   A super Array class in EasyASP
'##                 Support Array and Hashmap
'##
'######################################################################
Class EasyASP_List
  Public Size
  Private o_hash, o_map
  Private a_list
  Private i_count, i_comp
  Private s_separator
  
  Private Sub Class_Initialize
    Easp.Error("error-list-indexnull") = Easp.Lang("error-list-indexnull")
    Easp.Error("error-list-indexwrong") = Easp.Lang("error-list-indexwrong")
    Easp.Error("error-list-outofrange") = Easp.Lang("error-list-outofrange")
    Easp.Error("error-list-outofhash") = Easp.Lang("error-list-outofhash")
    Easp.Error("error-list-notlist") = Easp.Lang("error-list-notlist")
    Set o_hash = Server.CreateObject("Scripting.Dictionary")
    Set o_map  = Server.CreateObject("Scripting.Dictionary")
    Size = 0
    i_comp = 1
    s_separator = ""
    Set a_list = NewSimpleArray
  End Sub
  
  Private Sub Class_Terminate
    Set o_map  = Nothing
    Set o_hash = Nothing
  End Sub
  
  '建新Easp List类实例
  Public Function [New]()
    Set [New] = New EasyASP_List
    [New].IgnoreCase = Me.IgnoreCase
  End Function
  '建新Easp List数组实例
  Public Function NewArray(ByVal a)
    Set NewArray = New EasyASP_List
    NewArray.IgnoreCase = Me.IgnoreCase
    NewArray.Data = a
  End Function
  '建新Easp List Hash实例
  Public Function NewHash(ByVal a)
    Set NewHash = New EasyASP_List
    NewHash.IgnoreCase = Me.IgnoreCase
    NewHash.Hash = a
  End Function

  '新建Easp动态数组
  '说明：动态数组可以直接向任意下标赋值，用 .GetArray 取出数组
  Public Function NewSimpleArray()
    Set NewSimpleArray = New EasyASP_Json_Array
  End Function
  
  '是否忽略大小写
  Public Property Let IgnoreCase(ByVal b)
    i_comp = Easp.IIF(b, 1, 0)
  End Property
  Public Property Get IgnoreCase
    IgnoreCase = (i_comp = 1)
  End Property
  
  '设置和读取某一项值
  Public Property Let At(ByVal n, ByVal v)
    If Easp.IsN(n) Then
      If Easp.Debug Then
        Easp.Error.Raise "error-list-indexnull"
      End If
      Exit Property
    End If
    If Easp.Str.Test(n,"^\d+$") Then
      '如果是数字就直接添加到数组下标
      If n > [End] Then
        'ReDim Preserve a_list(n)
        Size = n + 1
      End If
      a_list(n) = v
    ElseIf Easp.Str.Test(n,"^[\w\./]+$") Then
    '如果是字符串
      If Not o_map.Exists(n) Then
      '如果散列中没有此项，添加映射关系及数组值
        o_map(n) = Size
        o_map(Size) = n
        Push v
      Else
      '如果已有该项，更新数组值
        a_list(o_map(n)) = v
      End If
    Else
      If Easp.Debug Then
        Easp.Error.Detail = n
        Easp.Error.Raise "error-list-indexwrong"
      End If
    End If
  End Property
  Public Default Property Get At(ByVal n)
    If Easp.Str.Test(n,"^\d+$") Then
      If n < Size Then
        At = a_list(n)
      Else
        At = Null
        If Easp.Debug Then
          Easp.Error.Detail = Array(n, [End])
          Easp.Error.Raise "error-list-outofrange"
        End If
      End If
    ElseIf Easp.Str.Test(n,"^[\w-\./]+$") Then
      If o_map.Exists(n) Then
        At = a_list(o_map(n))
      Else
        At = Null
        If Easp.Debug Then
          Easp.Error.Detail = Array(n, [End])
          Easp.Error.Raise "error-list-outofhash"
        End If
      End If
    End If
  End Property
  
  '设置和读取用普通数组方式设置初始值的分隔符
  Public Property Let Separator(ByVal string)
    s_separator = string
  End Property
  Public Property Get Separator
    Separator = s_separator
  End Property
  
  '设置源数组为普通数组或取出为普通数组
  Public Property Let Data(ByVal a)
    Data__ a, 0
  End Property
  Public Property Get Data
    Data = a_list.GetArray()
  End Property
  
  '设置源数组为哈希(Hash)表或取出为普通数组
  Public Property Let Hash(ByVal a)
    Data__ a, 1
  End Property
  Public Property Get Hash
    Dim arr, i
    arr = a_list.GetArray()
    For i = 0 To [End]
      If o_map.Exists(i) Then
        arr(i) = o_map(i) & ":" & arr(i)
      End If
    Next
    Hash = arr
  End Property
  '取值原型
  Public Sub Data__(ByVal a, ByVal t)
    Dim arr, i, j
    If isArray(a) Then
      a_list.SetArray a
      Size = a_list.Length
      If t = 0 Then Exit Sub
      For i = 0 To Ubound(a)
        If Instr(a(i),":")>0 Then
          j = Easp.Str.GetName(a(i),":")
          If Not o_map.Exists(j) Then
            'Easp.WN j
            o_map.Add i, j
            o_map.Add j, i
          End If
          a(i) = Easp.Str.GetValue(a(i),":")
        End If
      Next
      a_list.SetArray a
    Else
			arr = Easp.IIF(Easp.IsN(s_separator), Split(a), Split(a, s_separator))
			a_list.SetArray arr
			Size = a_list.Length
			If t = 0 Then Exit Sub
			'如果有Hash特征值
			If Instr(a, ":")>0 Then
				For i = 0 To Ubound(arr)
					'如果此元素是Hash下标
					If Instr(arr(i),":")>0 Then
						j = Easp.Str.GetName(arr(i),":")
						If Not o_map.Exists(j) Then
							o_map.Add i, j
							o_map.Add j, i
						End If
						arr(i) = Easp.Str.GetValue(arr(i),":")
					End If
				Next
			  a_list.SetArray arr
			End If
    End If
  End Sub
  
  '设置或读取Hash映射关系字典
  Public Property Let Maps(ByVal d)
    If TypeName(d) = "Dictionary" Then CloneDic__ o_map, d
  End Property
  Public Property Get Maps
    Set Maps = o_map
  End Property
  
  '返回数组的长度
  Public Property Get Length
    Length = Size
  End Property
  
  '返回数组的最大下标
  Public Property Get [End]
    [End] = Size - 1
  End Property
  
  '返回数组的有效长度（非空值）
  Public Property Get Count
    Dim i,j : j = 0
    For i = 0 To Size-1
      If Easp.Has(At(i)) Then j = j + 1
    Next
    Count = j
  End Property
  
  '返回数组的第一个元素值
  Public Property Get First
    First = At(0)
  End Property
  
  '返回数组的最后一个元素值
  Public Property Get Last
    Last = At([End])
  End Property
  
  '返回数组中的最大值
  Public Property Get Max
    Dim i, v
    v = At(0)
    If Size > 1 Then
      For i = 1 To [End]
        If Compare__("gt", At(i), v) Then v = At(i)
      Next
    End If
    Max = v
  End Property
  
  '返回数组中的最小值
  Public Property Get Min
    Dim i, v
    v = At(0)
    If Size > 1 Then
      For i = 1 To [End]
        If Compare__("lt", At(i), v) Then v = At(i)
      Next
    End If
    Min = v
  End Property
  
  '序列化Hash表，即转化为a=1&b=2&c=3型字符串
  Public Property Get Serialize
    Dim tmp, i : tmp = ""
    For i = 0 To [End]
      If o_map.Exists(i) Then
        tmp = tmp & "&" & o_map(i) & "=" & Server.URLEncode(At(i))
      End If
    Next
    If Len(tmp)>1 Then tmp = Mid(tmp,2)
    Serialize = tmp
  End Property
  
  '检测是否包含某个下标
  Public Function HasIndex(ByVal i)
    HasIndex = Index(i) >= 0
  End Function
  
  '返回Hash名称所在的下标数字
  Public Function Index(ByVal i)
    If isNumeric(i) Then
      Index = Easp.IIF(i >= 0 And i <= [End], i, -1)
    Else
      If o_map.Exists(i) Then
        Index = o_map(i)
      Else
        Index = -1
      End If
    End If
  End Function
  
  '返回数字下标的Hash名称
  Public Function IndexHash(ByVal i)
    If isNumeric(i) Then
      IndexHash = Easp.IfThen(o_map.Exists(i), o_map(i))
    Else
      IndexHash = Easp.IfThen(o_map.Exists(i), i)
    End If
    IndexHash = Easp.IfHas(IndexHash, i)
  End Function
  
  '比较函数
  '参数： @t - 比较方式： lt 小于
  Private Function Compare__(ByVal t, ByVal a, ByVal b)
    Dim isStr : isStr = False
    If VarType(a) = 8 Or VarType(b) = 8 Then
      isStr = True
      If IsNumeric(a) And IsNumeric(b) Then isStr = False
      If IsDate(a) And IsDate(b) Then isStr = False
    End If
    If isStr Then
      Select Case LCase(t)
        Case "lt" Compare__ = (StrComp(a,b,i_comp) = -1)
        Case "gt" Compare__ = (StrComp(a,b,i_comp) = 1)
        Case "eq" Compare__ = (StrComp(a,b,i_comp) = 0)
        Case "lte" Compare__ = (StrComp(a,b,i_comp) = -1 Or StrComp(a,b,i_comp) = 0)
        Case "gte" Compare__ = (StrComp(a,b,i_comp) = 1 Or StrComp(a,b,i_comp) = 0)
      End Select
    Else
      Select Case LCase(t)
        Case "lt" Compare__ = (a < b)
        Case "gt" Compare__ = (a > b)
        Case "eq" Compare__ = (a = b)
        Case "lte" Compare__ = (a <= b)
        Case "gte" Compare__ = (a >= b)
      End Select
    End If
  End Function
  
  '添加一个元素到数组开头
  Public Sub UnShift(ByVal v)
    Insert 0, v
  End Sub
  '添加一个元素到数组开头并返回新数组对象
  Public Function UnShift_(ByVal v)
    Set UnShift_ = Me.Clone
    UnShift_.UnShift v
  End Function
  
  '删除数组第一个元素
  Public Sub Shift
    [Delete] 0
  End Sub
  '删除数组第一个元素并返回新数组对象
  Public Function Shift_
    Set Shift_ = Me.Clone
    Shift_.Shift
  End Function
  
  '添加一个元素到数组结尾
  Public Sub Push(ByVal v)
    'ReDim Preserve a_list(Size)
    a_list(Size) = v
    Size = Size + 1
  End Sub
  '添加一个元素到数组结尾并返回新数组对象
  Public Function Push_(ByVal v)
    Set Push_ = Me.Clone
    Push_.Push v
  End Function
  
  '删除数组最后一个元素
  Public Sub Pop
    RemoveMap__ [End]
    'ReDim Preserve a_list([End]-1)
    a_list.Remove([End])
    Size = Size - 1
  End Sub
  '删除数组最后一个元素并返回新数组对象
  Public Function Pop_
    Set Pop_ = Me.Clone
    Pop_.Pop
  End Function
  Private Sub RemoveMap__(ByVal i)
    If o_map.Exists(i) Then
      o_map.Remove o_map(i)
      'Easp.WN "=Delete==mapRemove:" & o_map(i)
      o_map.Remove i
      'Easp.WN "=Delete==mapRemove:" & i
    End If
  End Sub
  Private Sub UpFrom__(ByVal n, ByVal i)
    If n = i Then Exit Sub
    If o_map.Exists(i) Then
      o_map(o_map(i)) = n
      o_map(n) = o_map(i)
      o_map.Remove i
      'Easp.WN "=Delete==UpFromRemove:" & i & "  o_map(count):" & o_map.count
    End If
    At(n) = At(i)
  End Sub
  
  '在指定下标插入一个元素或一个数组
  Public Sub Insert(ByVal n, ByVal v)
    Dim i,j
    If n > [End] Then
    '如果下标大于最大下标
      If isArray(v) Then
      '如果插入一个数组，逐个赋值
        For i = 0 To UBound(v)
          At(n+i) = v(i)
        Next
      Else
      '是字符串直接赋值
        At(n) = v
      End If
    Else
    '如果插入到数组中间
      '如果插入一个数组
      For i = Size To (n+1) Step -1
      '将原数组插入点之后的值移动到新位置（腾出位置）
        If isArray(v) Then
        '如果是数组，要腾出数组的长度个位置
          UpFrom__ i+UBound(v), i-1
          'Easp.WN "把 " &i-1& "的值修改到 " &i+UBound(v)& " 上"
        Else
        '否则只腾出一个位置
          UpFrom__ i, i-1
        End If
      Next
      '把新值插入腾出的位置上
      If isArray(v) Then
        For i = 0 To UBound(v)
          At(n+i) = v(i)
        Next
      Else
        At(n) = v
      End If
    End If
  End Sub
  '在指定下标插入一个元素或一个数组并返回新数组对象
  Public Function Insert_(ByVal n, ByVal v)
    Set Insert_ = Me.Clone
    Insert_.Insert n, v
  End Function
  
  '检测数组中是否包含某个值
  Public Function Has(ByVal v)
    Has = (indexOf__(a_list.GetArray, v) > -1)
  End Function
  
  '检测某个值在数组中的下标
  Public Function IndexOf(ByVal v)
    IndexOf = indexOf__(a_list.GetArray, v)
  End Function
  '检测某个值在数组中的Hash名称
  Public Function IndexOfHash(ByVal v)
    Dim i : i = indexOf__(a_list.GetArray, v)
    If i = -1 Then IndexOfHash = Empty : Exit Function
    If o_map.Exists(i) Then
      IndexOfHash = o_map(i)
    Else
      IndexOfHash = Empty
    End If
  End Function
  Private Function indexOf__(ByVal arr, ByVal v)
    Dim i
    indexOf__ = -1
    For i = 0 To UBound(arr)
      If Compare__("eq", arr(i),v) Then
        indexOf__ = i
        Exit For
      End If
    Next
  End Function
  
  '删除一个或多个元素
  Public Sub [Delete](ByVal n)
    Dim tmp,a,x,y,i
    If Instr(n, ",")>0 Or Instr(n,"-")>0 Then
    '如果是删除多个元素
      '开始符号
      n = Replace(n,"\s","0")
      '结束符号
      n = Replace(n,"\e",[End])
      a = Split(n, ",")
      For i = 0 To Ubound(a)
      '单独处理每项
        If i>0 Then tmp = tmp & ","
        '如果有区间
        If Instr(a(i),"-")>0 Then
          '取首元素
          x = Trim(Easp.Str.GetName(a(i),"-"))
          '取尾元素
          y = Trim(Easp.Str.GetValue(a(i),"-"))
          '如果是Hash则取出为数字下标
          If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
          If Not Isnumeric(y) And o_map.Exists(y) Then y = o_map(y)
          '重新组合为数字区间
          tmp = tmp & x & "-" & y
        Else
        '如果是单项
          x = Trim(a(i))
          '如果是Hash则取出为数字下标
          If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
          tmp = tmp & x
        End If
      Next
      '将要删除的编号组转换为数组并排序
      a = Split(tmp,",")
      a = SortArray(a,0,UBound(a))
      tmp = "0-"
      For i = 0 To Ubound(a)
        If Instr(a(i),"-")>0 Then
          x = Easp.Str.GetName(a(i),"-")
          y = Easp.Str.GetValue(a(i),"-")
          'Easp.WN a(i)
          tmp = tmp & x-1 & ","
          tmp = tmp & y+1 & "-"
        Else
          tmp = tmp & a(i)-1 & "," & a(i)+1 & "-"
        End If
      Next
      tmp = tmp & [End]
      'Easp.WN tmp
      Slice tmp
    Else
    '只删除一项
      If Not isNumeric(n) And o_map.Exists(n) Then
        n = o_map(n)
        RemoveMap__ n
      End If
      For i = n+1 To [End]
        UpFrom__ i-1, i
      Next
      Pop
    End If
  End Sub
  '删除一个或多个元素并返回新数组对象
  Public Function Delete_(ByVal n)
    Set Delete_ = Me.Clone
    Delete_.Delete n
  End Function

  '移除数组中的重复元素只保留一个
  Public Sub Uniq()
    Dim arr(),i,j : j = 0
    ReDim arr(-1)
    If o_hash.Count>0 Then o_hash.RemoveAll
    For i = 0 To [End]
      '如果新数组中没有该值
      If indexOf__(arr, At(i)) = -1 Then
        ReDim Preserve arr(j)
        arr(j) = At(i)
        'Easp.WN "把元素" & i & "存入了新数组的 " & j
        If o_map.Exists(i) Then
          o_hash.Add j, o_map(i)
          o_hash.Add o_map(i), j
          'Easp.WN "把Hash中的第("&i&")项" &o_map(i)& "存入新Hash的" & j
        End If
        j = j + 1
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '移除数组中的重复元素只保留一个并返回新数组对象
  Public Function Uniq_()
    Set Uniq_ = Me.Clone
    Uniq_.Uniq
  End Function
  Private Sub CloneDic__(ByRef map, ByRef hash)
    Dim key
    If map.Count > 0 Then map.RemoveAll
    For Each key In hash
      map(key) = hash(key)
    Next
  End Sub
  
  '让数组随机排序(洗牌)
  Public Sub Rand
    Dim i, j, tmp, Ei, Ej, Ti, Tj
    For i = 0 To [End]
      j = Easp.Str.RandomNumber(0,[End])
      '检测是否为Hash，如果是Hash就把值存起来
      Ei = o_map.Exists(i)
      Ej = o_map.Exists(j)
      If Ei Then Ti = o_map(i)
      If Ej Then Tj = o_map(j)
      '数组值互换
      tmp = At(j)
      At(j) = At(i)
      At(i) = tmp
      'Hash值互换
      If Ei Then
        o_map(j) = Ti
        o_map(Ti) = j
      End If
      If Ej Then
        o_map(i) = Tj
        o_map(Tj) = i
      End If
      '如果其中至少一个为空，则删除在Hash中的此下标
      If Not (Ei And Ej) Then
        If Ei Then o_map.Remove i
        If Ej then o_map.Remove j
      End If
    Next
  End Sub
  '让数组随机排序并返回新数组对象
  Public Function Rand_()
    Set Rand_ = Me.Clone
    Rand_.Rand
  End Function
  
  '将数组倒序排列
  Public Sub Reverse
    Dim arr(),i,j : j = 0
    ReDim arr([End])
    If o_hash.Count>0 Then o_hash.RemoveAll
    For i = [End] To 0 Step -1
      arr(j) = At(i)
      If o_map.Exists(i) Then
        o_hash.Add j, o_map(i)
        o_hash.Add o_map(i), j
      End If
      j = j + 1
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '将数组倒序排列并返回新数组对象
  Public Function Reverse_()
    Set Reverse_ = Me.Clone
    Reverse_.Reverse
  End Function

  '搜索包含指定字符串的元素
  Public Sub Search(ByVal s)
    Search__ s, True
  End Sub
  '搜索包含指定字符串的元素并返回新数组对象
  Public Function Search_(ByVal s)
    Set Search_ = Me.Clone
    Search_.Search s
  End Function

  '搜索不包含指定字符串的元素
  Public Sub SearchNot(ByVal s)
    Search__ s, False
  End Sub
  '搜索不包含指定字符串的元素并返回新数组对象
  Public Function SearchNot_(ByVal s)
    Set SearchNot_ = Me.Clone
    SearchNot_.SearchNot s
  End Function
  
  Private Sub Search__(ByVal s, ByVal keep)
    Dim arr,i,tmp
    '搜索结果
    arr = Filter(a_list.GetArray, s, keep, i_comp)
    If o_map.Count = 0 Then
      Data = arr
    Else
      AddHash__ arr
    End If
  End Sub
  
  '删除数组中的空元素
  Public Sub Compact
    Dim arr(), i, j : j = 0
    If o_hash.Count>0 Then o_hash.RemoveAll
    For i = 0 To [End]
      If Easp.Has(At(i)) Then
        ReDim Preserve arr(j)
        arr(j) = At(i)
        If o_map.Exists(i) Then
          o_hash.Add j, o_map(i)
          o_hash.Add o_map(i), j
        End If
        j = j + 1
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '删除数组中的空元素并返回新数组对象
  Public Function Compact_()
    Set Compact_ = Me.Clone
    Compact_.Compact
  End Function
  
  '清空数组
  Public Sub Clear
    a_list.Clear
    If o_map.Count>0 Then o_map.RemoveAll
    Size = 0
  End Sub
  
  '数组排序
  Public Sub Sort
    Dim arr
    arr = a_list.GetArray
    arr = SortArray(arr, 0, [End])
    If o_map.Count = 0 Then
      Data = arr
    Else
      AddHash__ arr
    End If
  End Sub
  '数组排序并返回新数组对象
  Public Function Sort_()
    Set Sort_ = Me.Clone
    Sort_.Sort
  End Function
  Private Function SortArray(ByRef arr, ByRef low, ByRef high)
    If Not IsArray(arr) Then Exit Function
    If Easp.IsN(arr) Then Exit Function
    Dim l, h, m, v, x
    l = low : h = high
    m = (low + high) \ 2 : v = arr(m)
    Do While (l <= h)
      Do While (Compare__("lt",arr(l),v) And l < high)
      'Do While (arr(l) < v And l < high)
        'Easp.WN arr(l) & " &lt; " & v
        l = l + 1
      Loop
      Do While (Compare__("lt",v,arr(h)) And h > low)
      'Do While (v < arr(h) And h > low)
        'Easp.WN v & " &lt; " & arr(h)
        h = h - 1
      Loop
      If l <= h Then
        x = arr(l) : arr(l) = arr(h) : arr(h) = x   
        l = l + 1 : h = h - 1         
      End If
    Loop
    If (low < h) Then arr = SortArray(arr, low, h)
    If (l < high) Then arr = SortArray(arr,l, high)
    SortArray = arr
  End Function
  'For Sort & Search & SearchNot
  Private Sub AddHash__(ByVal arr)
    Dim tmp, i
    If o_hash.Count > 0 Then o_hash.RemoveAll
    For i = 0 To Ubound(arr)
      '如果结果中有Hash下标
      'Easp.WN arr(i) & " index:" & IndexOfHash(arr(i))
      If IndexOfHash(arr(i))>"" Then
        '取出这个下标
        tmp = IndexOfHash(arr(i))
        '添加到新的索引值（只添加一次）
        If Not o_hash.Exists(tmp) Then 
          o_hash.Add i, tmp
          o_hash.Add tmp, i
        End If
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  
  '按下标取出部分元素而删除数组的其它元素
  Public Sub Slice(ByVal s)
    Dim a,i,j,k,x,y,arr
    If o_hash.Count>0 Then o_hash.RemoveAll
    s = Replace(s,"\s",0)
    s = Replace(s,"\e",[End])
    a = Split(s, ",")
    arr = Array() : k = 0
    For i = 0 To Ubound(a)
      'Easp.WN "Big:" & k
      '如果是区间
      If Instr(a(i),"-")>0 Then
        '如果是Hash则取出为数字下标
        x = Trim(Easp.Str.GetName(a(i),"-"))
        y = Trim(Easp.Str.GetValue(a(i),"-"))
        If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
        If Not Isnumeric(y) And o_map.Exists(y) Then y = o_map(y)
        x = Int(x) : y = Int(y)
        'Easp.WN x & "-" & y
        For j = x To y
          ReDim Preserve arr(k)
          'Easp.WN "Small:"&k & "=" & x & "-" & y
          arr(k) = At(j)
          If o_map.Exists(j) Then
            'Easp.WN o_map(j) & " "&k&" " & o_hash.Exists(o_map(j))
            '如果出现多个相同的Hash下标则只保留第一个Hash下标
            If Not o_hash.Exists(o_map(j)) Then
              o_hash.Add k, o_map(j)
              o_hash.Add o_map(j), k
            End If
          End If
          k = k + 1
        Next
      Else
        ReDim Preserve arr(k)
        x = Trim(a(i))
        'Easp.WN x
        If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
        x = Int(x)
        If o_map.Exists(x) Then
          '如果出现多个相同的Hash下标则只保留第一个Hash下标
          'Easp.WN o_map(x) & " " & o_hash.Exists(o_map(x))
          If Not o_hash.Exists(o_map(x)) Then
            o_hash.Add k, o_map(x)
            o_hash.Add o_map(x), k
          End If
        End If
        arr(k) = At(x)
        k = k + 1
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '按下标取出部分元素并返回新数组对象
  Public Function Slice_(ByVal s)
    Set Slice_ = Me.Clone
    Slice_.Slice s
  End Function
  '按下标取出部分元素并返回新数组对象
  Public Function [Get](ByVal s)
    Set [Get] = Slice_(s)
  End Function
  
  '返回将数组元素用字符连接后的字符串
  Public Function Join(ByVal s)
    Dim j
    Set j = New EasyASP_List_Join
    Join = j.J(a_list.GetArray, s)
    Set j = Nothing
  End Function
  '返回将数组元素用字符连接后的字符串并返回新数组对象
  Public Function Join_(ByVal s)
    Set Join_ = Me.Clone
    Join_.Join s
  End Function
  
  '将数组转换成用逗号隔开的字符串
  Public Function ToString()
    ToString = Join(",")
  End Function
  
  '取出为普通数组(无Hash标识的普通数组)
  Public Function ToArray
    ToArray = a_list.GetArray
  End Function
  
  '复制List对象
  Public Function Clone
    Set Clone = Me.New
    Clone.Data = a_list.GetArray
    If o_map.Count>0 Then Clone.Maps = o_map
  End Function
  
  '=============
  '=以下是循环处理部分
  '=============


  '按元素值进行循环处理并返回新值到数组
  Public Sub Map(ByVal f)
    '意思是依次对数组中的元素调用某个方法进行处理并将处理后的值替换到数组
    Map__ f, 0
  End Sub
  '按元素值进行循环处理并返回新数组对象
  Public Function Map_(ByVal f)
    Set Map_ = Me.Clone
    Map_.Map f
  End Function
  
  '按元素值进行循环调用方法
  Public Sub [Each](ByVal f)
    '意思是依次把数组中的元素作用参数调用某个方法
    Map__ f, 1
  End Sub
  Private Sub Map__(ByVal f, ByVal t)
    Dim i, tmp
    For i = 0 To [End]
      tmp = Value__(At(i))
      If t = 0 Then
        '返回值到数组
        If Instr(f, "%i") > 0 Then
          'Easp.WN Replace(f, "%i", tmp)
          At(i) = Eval(Replace(f, "%i", tmp))
        Else
          At(i) = Eval(f & "("& tmp &")")
        End If
      ElseIf t = 1 Then
        '直接执行
        If Instr(f, "%i") > 0 Then
          ExecuteGlobal Replace(f, "%i", tmp)
        Else
          ExecuteGlobal f & "("& tmp &")"
        End If
      End If
    Next
  End Sub
  Private Function Value__(ByVal s)
    Dim tmp
    Select Case VarType(s)
      Case 7,8 tmp = """" & s & """"
      Case Else tmp = s
    End Select
    Value__ = tmp
  End Function
  
  '返回第一个符合表达式的元素值
  Public Function Find(ByVal f)
    Dim i, k, tmp
    '默认标识符为 i
    k = "i"
    If Easp.Str.Test(f,"[a-zA-Z]+:(.+)") Then
      '如果有自定义的标识符
      k = Easp.Str.GetName(f,":")
      f = Easp.Str.GetValue(f,":")
    End If
    k = "%" & k
    For i = 0 To [End]
      tmp = Replace(Trim(f), k, Value__(At(i)))
      If Eval(tmp) Then
        Find = At(i) : Exit Function
      End If
    Next
    Find = Empty
  End Function
  
  '删除所有不符合表达式条件的元素
  Public Sub [Select](ByVal f)
    Select__ f, 0
  End Sub
  '用所有符合表达式条件的元素组成新数组对象
  Public Function Select_(ByVal f)
    Set Select_ = Me.Clone
    Select_.Select f
  End Function
  Private Sub Select__(ByVal f, ByVal t)
    Dim i, j, k, tmp, arr
    arr = Array() : j = 0
    If o_hash.Count>0 Then o_hash.RemoveAll
    k = "i"
    If Easp.Str.Test(f,"[a-zA-Z]+:(.+)") Then
      k = Easp.Str.GetName(f,":")
      f = Easp.Str.GetValue(f,":")
    End If
    k = "%" & k
    For i = 0 To [End]
      tmp = Replace(Trim(f), k, Value__(At(i)))
      If t = 0 Then
        tmp = Eval(tmp)
      ElseIf t = 1 Then
        tmp = (Not Eval(tmp))
      End If
      If tmp Then
        ReDim Preserve arr(j)
        arr(j) = At(i)
        If o_map.Exists(i) Then
          o_hash.Add j, o_map(i)
          o_hash.Add o_map(i), j
        End If
        j = j + 1
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  
  '删除所有符合表达式条件的元素
  Public Sub Reject(ByVal f)
    Select__ f, 1
  End Sub
  '用所有不符合表达式条件的元素组成新数组对象
  Public Function Reject_(ByVal f)
    Set Reject_ = Me.Clone
    Reject_.Reject f
  End Function
  
  '按元素值返回符合正则表达式的元素
  Public Sub Grep(ByVal g)
    Dim i,j,arr
    arr = Array() : j = 0
    If o_hash.Count>0 Then o_hash.RemoveAll
    For i = 0 To [End]
      If Easp.Str.Test(At(i),g) Then
        ReDim Preserve arr(j)
        arr(j) = At(i)
        If o_map.Exists(i) Then
          o_hash.Add j, o_map(i)
          o_hash.Add o_map(i), j
        End If
        j = j + 1
      End If
    Next
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '选择符合正则表达式的元素并返回新数组对象
  Public Function Grep_(ByVal g)
    Set Grep_ = Me.Clone
    Grep_.Grep g
  End Function
  
  '按元素值进行循环处理后并排序
  Public Sub SortBy(ByVal f)
    Map f : Sort
  End Sub
  '按元素值进行循环处理后排序并返回新数组对象
  Public Function SortBy_(ByVal f)
    Set SortBy_ = Me.Clone
    SortBy_.SortBy f
  End Function
  
  '=============
  '以下是数组运算处理部分
  '=============

  '把一个数组重复多次
  Public Sub Times(ByVal t)
    Dim i, arr
    arr = a_list.GetArray
    For i = 1 To t
      Insert Size, arr
    Next
  End Sub
  '把一个数组重复多次并返回新数组对象
  Public Function Times_(ByVal t)
    Set Times_ = Me.Clone
    Times_.Times t
  End Function
  '判断是不是Easp的List对象
  Private Function IsList(ByVal o)
    IsList = Easp.Str.IsSame(TypeName(o), "EasyASP_list")
  End Function
  
  '附加数组

  '把一个数组拼接到另一个数组最后
  Public Sub Splice(ByVal o)
    If Not isArray(o) And Not isList(o) Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "List.Splice"
        Easp.Error.Raise "error-list-notlist"
      End If
      Exit Sub
    End If
    Dim omap,dic,i
    '如果是数组，直接拼接在最后
    If isArray(o) Then
      Insert Size, o
    '如果是List对象
    ElseIf IsList(o) Then
      '先检测是否有Hash值
      Set omap = o.Maps
      '如果有Hash值
      If omap.Count > 0 Then
        For i = 0 To o.End
          'Easp.WN Easp.Str("{3}...{1} = {2}", Array(omap.Exists(i), (Not o_map.Exists(omap(i))), i))
          '取出Hash值名，存入原数组
          If omap.Exists(i) And (Not o_map.Exists(omap(i))) Then
            o_map.Add Size + i, omap(i)
            o_map.Add omap(i), Size + i
            'Easp.WN Easp.Str("{1} = {2}", Array(omap(i), Size + i))
          End If
        Next
      End If
      '把新值插入原数组
      Insert Size, o.Data
      Set omap = Nothing
    End If
  End Sub
  '把一个数组拼接到另一个数组最后并返回新数组对象
  Public Function Splice_(ByVal o)
    Set Splice_ = Me.Clone
    Splice_.Splice o
  End Function
  
  '数组合集
  '把两个数组合并并删除重复项
  Public Sub Merge(ByVal o)
    Splice o
    Uniq
  End Sub
  '把两个数组合并并返回新数组对象
  Public Function Merge_(ByVal o)
    Set Merge_ = Me.Clone
    Merge_.Merge o
  End Function
  
  '数组交集
  '取出在两个数组中都存在的元素
  Public Sub Inter(ByVal o)
    If Not isArray(o) And Not isList(o) Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "List.Inter"
        Easp.Error.Raise "error-list-notlist"
      End If
      Exit Sub
    End If
    Dim i,j,k,omap,arr
    arr = Array() : j = 0
    If o_hash.Count>0 Then o_hash.RemoveAll
    '如果是数组
    If isArray(o) Then
      '遍历数组
      For i = 0 To Ubound(o)
        '如果数组中的值在List中
        If Has(o(i)) Then
          ReDim Preserve arr(j)
          '把值存入临时数组
          arr(j) = o(i)
          '取值在原List中的下标
          k = IndexOf(o(i))
          '如果是hash值，则写入中转字典
          If o_map.Exists(k) Then
            o_hash.Add j, o_map(k)
            o_hash.Add o_map(k), j
          End If
          j = j + 1
        End If
      Next
    '如果是List对象
    ElseIf IsList(o) Then
      '取出Hash映射表
      Set omap = o.Maps
      '遍历List对象
      For i = 0 To o.End
        '如果在原List中存在
        If Has(o(i)) Then
          ReDim Preserve arr(j)
          arr(j) = o(i)
          k = IndexOf(o(i))
          '检测在原List中是否是Hash值
          If o_map.Exists(k) Then
            o_hash.Add j, o_map(k)
            o_hash.Add o_map(k), j
          '检测在新List中是否是Hash值
          ElseIf omap.Exists(i) Then
            o_hash.Add j, omap(i)
            o_hash.Add omap(i), j
          End If
          j = j + 1
        End If
      Next
      Set omap = Nothing
    End If
    '把新值存入当前List
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '取出在两个数组中都存在的元素并返回新数组对象
  Public Function Inter_(ByVal o)
    Set Inter_ = Me.Clone
    Inter_.Inter o
  End Function
  
  '数组差集
  '取出在一个数组中存在而在另一个数组中不存在的元素
  Public Sub Diff(ByVal o)
    If Not isArray(o) And Not isList(o) Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "List.Diff"
        Easp.Error.Raise "error-list-notlist"
      End If
      Exit Sub
    End If
    Dim i,j,arr,a
    arr = Array() : j = 0
    If o_hash.Count>0 Then o_hash.RemoveAll
    If isArray(o) Then
      a = o
      Set o = Me.New
      o.Data = a
    End If
    For i = 0 To [End]
      If Not o.Has(At(i)) Then
        ReDim Preserve arr(j)
        arr(j) = At(i)
        If o_map.Exists(i) Then
          o_hash.Add j, o_map(i)
          o_hash.Add o_map(i), j
        End If
        j = j + 1
      End If
    Next
    '把新值存入当前List
    Data = arr
    CloneDic__ o_map, o_hash
    o_hash.RemoveAll
  End Sub
  '取数组差集并返回新数组对象
  Public Function Diff_(ByVal o)
    Set Diff_ = Me.Clone
    Diff_.Diff o
  End Function
  
  '比较数组
  '比较两个数组的大小
  Public Function Eq(ByVal o)
    If Not isArray(o) And Not isList(o) Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "List.Eq"
        Easp.Error.Raise "error-list-notlist"
      End If
      Exit Function
    End If
    Dim a, e, m, i, j
    If isArray(o) Then
      a = o
      e = Ubound(a)
    ElseIf isList(o) Then
      a = o.Data
      e = o.End
    End If
    m = Easp.IIF([End] < e, [End], e)
    '遍历List，逐个比较元素，如不同则出结果
    For i = 0 To m
      If Compare__("gt", At(i), a(i)) Then
        Eq = 1 : Exit Function
      ElseIf Compare__("lt", At(i), a(i)) Then
        Eq = -1 : Exit Function
      End If
    Next
    '如果在短的数组遍历完后仍相等，则长的那个数组较大
    If [End] > e Then
      Eq = 1
    ElseIf [End] < e Then
      Eq = -1
    Else
      Eq = 0
    End If
  End Function
  
  '检测一个数组是否是本数组的子集
  Public Function Son(ByVal o)
    If Not isArray(o) And Not isList(o) Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "List.Son"
        Easp.Error.Raise "error-list-notlist"
      End If
      Exit Function
    End If
    Son = True
    Dim i
    If isList(o) Then o = o.Data
    For i = 0 To Ubound(o)
      If Not Has(o(i)) Then Son = False : Exit Function
    Next
  End Function
End Class
Class EasyASP_List_Join
  Function J(ByRef arr, ByVal separator)
    J = Join(arr, separator)
  End Function
End Class
%>