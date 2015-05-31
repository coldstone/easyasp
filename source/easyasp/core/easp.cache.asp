<%
'######################################################################
'## easp.cache.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Cache Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com) & SunYu
'## Update Date :   2014-05-01 23:54:53
'## Description :   Save and Get Cache With EasyASP
'##
'######################################################################
Class EasyASP_Cache
  Public Items, CountEnabled, Expires, FileType
  Private s_path, b_fsoOn
  '构造函数
  Private Sub Class_Initialize
    Set Items = Server.CreateObject("Scripting.Dictionary")
    Items.CompareMode = 1
    s_path = Server.MapPath("/_cache") & "\"
    CountEnabled = True
    Expires = 5
    FileType = ".easpcache"
    Easp.Error("error-cache-notfound") = Easp.Lang("error-cache-notfound")
    Easp.Error("error-cache-invalid-object") = Easp.Lang("error-cache-invalid-object")
    Easp.Error("error-cache-invalid-file") = Easp.Lang("error-cache-invalid-file")
  End Sub
  '析构函数
  Private Sub Class_Terminate
    Set Items = Nothing
  End Sub
  '建新Easp缓存类实例
  Public Function [New]()
    Set [New] = New EasyASP_Cache
  End Function
  '取当前所有缓存数量
  Public Property Get Count
    Count = Easp.IIF(CountEnabled,Easp_Cache_Count,-1)
  End Property
  '添加缓存值
  Public Property Let Item(ByVal p, ByVal v)
    If IsNull(p) Then p = ""
    If Not IsObject(Items(p)) Then
      Set Items(p) = New Easp_Cache_Info
      Items(p).CountEnabled = CountEnabled
      Items(p).Expires = Expires
      Items(p).FileType = FileType
    End If
    Items(p).Name = p
    Items(p).Value = v
    Items(p).SavePath = s_path
  End Property
  '获取缓存值
  Public Default Property Get Item(ByVal p)
    If Not IsObject(Items(p)) Then
      Set Items(p) = New Easp_Cache_Info
      Items(p).Name = p
      Items(p).SavePath = s_path
      Items(p).CountEnabled = CountEnabled
      Items(p).Expires = Expires
      Items(p).FileType = FileType
    End If
    set Item = Items(p)
  End Property
  '设置文件缓存保存目录路径
  Public Property Let SavePath(ByVal s)
    If Not Instr(s,":") = 2 Then s = Server.MapPath(s)
    If Right(s,1) <> "\" Then s = s & "\"
    s_path = s
  End Property
  Public Property Get SavePath()
    SavePath = s_path
  End Property
  '保存所有文件缓存
  Public Sub SaveAll
    Dim f
    For Each f In Items
      Items(f).Save
    Next
  End Sub
  '保存所有内存缓存
  Public Sub SaveAppAll  
    Dim f 
    For Each f In Items
      Items(f).SaveApp
    Next
  End Sub
  '清除所有文件缓存
  Public Sub RemoveAll
    Dim f
    For Each f In Items
      Items(f).Remove
    Next
  End Sub
  '清除所有内存缓存
  Public Sub RemoveAppAll  
    Dim f 
    For Each f In Items
      Items(f).RemoveApp
    Next
  End Sub
  '清空缓存
  Public Sub [Clear]
    RemoveAll
    RemoveAppAll
    Easp.RemoveApplication "Easp_Cache_Count"
  End Sub
End Class
'统计缓存数量
Private Function Easp_Cache_Count()
  Easp_Cache_Count = 0
  Dim n : n = Easp.GetApplication("Easp_Cache_Count")
  If IsArray(n) Then
    If Ubound(n) = 1 Then Easp_Cache_Count = n(0)
  End If
End Function
'缓存计数更改
Private Function Easp_CacheCount_Change(ByVal a, ByVal t)
  Dim n : n = Easp.GetApplication("Easp_Cache_Count")
  If isArray(n) Then
    If Ubound(n) = 1 Then
      If TypeName(n(1)) = "Dictionary" Then
        If t = 1 Then n(1)(a) = a
        If t = -1 Then
          If n(1).Exists(a) Then n(1).Remove(a)
        End If
        Easp.SetApplication "Easp_Cache_Count", Array(n(1).Count,n(1))
      End If
    End If
  Else
    Dim dic : Set dic = Server.CreateObject("Scripting.Dictionary")
    If t = 1 Then dic(a) = a
    Easp.SetApplication "Easp_Cache_Count", Array(Easp.IIF(t=1,1,0),dic)
  End If
End Function
'缓存项处理方法
class Easp_Cache_Info
  Public SavePath, [Name], CountEnabled, FileType
  Private i_exp, d_exp, o_value
  Private Sub Class_Initialize
    i_exp = 5
    d_exp = ""
  End Sub
  Private Sub Class_Terminate
    If IsObject(o_value) Then Set o_value = Nothing
  End Sub
  '设置和读取缓存过期时间
  Public Property Let Expires(ByVal i)
    If isDate(i) Then
      '具体日期时间
      d_exp = CDate(i)
    ElseIf isNumeric(i) Then
      '数值（分钟）
      If i>0 Then
        i_exp = i
      ElseIf i=0 Then
        i_exp = 60*24*365*99
      End If
    End If
  End Property
  Public Property Get Expires()
    Expires = Easp.IfHas(d_exp, i_exp)
  End Property
  '设置和读取缓存的值
  Public Property Let [Value](ByVal s)
    If IsObject(s) Then
      Select Case TypeName(s)
        Case "Recordset"
        '如果是记录集
          Set o_value = s.Clone
        Case Else
        '如果是其它对象
          Set o_value = s
      End Select
    Else
      '其它直接赋值
      o_value = s
    End If
  End Property
  Public Default Property Get [Value]()
    '在内存缓存中取值
    Dim app : app = Easp.GetApplication(Me.Name)
    If IsArray(app) Then
      If UBound(app) = 1 Then
        If IsDate(app(0)) Then
          If IsObject(app(1)) Then
            Set [Value] = app(1)
            Exit Property
          Else
            [Value] = app(1)
            If Easp.Has([Value]) Then Exit Property
          End If
        End If
      End If
    End If
    '如果内存缓存中没有该值则在文件缓存中取
    If Easp.Fso.IsFile(FilePath) Then
      On Error Resume Next
      Dim rs
      set rs = Server.CreateObject("Adodb.Recordset")
      rs.Open FilePath
      If Err.Number <> 0 Then
        Err.Clear
        [Value] = Easp.Fso.Read(FilePath)
      Else
        Set [Value] = rs
      End If
    Else
      Easp.Error.FunctionName = "Cache:Item.Get"
      Easp.Error.Detail = Easp.Str.HtmlEncode(Me.Name)
      Easp.Error.Raise "error-cache-notfound"
    End If
  End Property
  '保存到内存缓存
  Public Sub SaveApp
    Dim appArr(1) : appArr(0) = Now()
    If IsObject(o_value) Then
      '保存字典对象和记录对象（记录集对象将自动转为二维数组）
      Select Case TypeName(o_value)
        Case "Dictionary"
          Set appArr(1) = o_value
        Case "Recordset"
          appArr(1) = o_value.GetRows(-1)
        Case Else
          Easp.Error.FunctionName = "Cache:Item.SaveApp"
          Easp.Error.Detail = Easp.Str.HtmlEncode(Me.Name)&" &gt; "&TypeName(o_value)
          Easp.Error.Raise "error-cache-invalid-object"
      End Select
    Else
      appArr(1) = o_value
    End If
    Easp.SetApplication Me.Name, appArr
    If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
  End Sub
  '保存到文件缓存
  Public Sub Save
    Select Case TypeName(o_value)
      Case "Recordset"
        Easp.Fso.CreateFile FilePath, "rs"
        Easp.Fso.DelFile FilePath
        o_value.Save FilePath, 1
        If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
      Case "String"
        Easp.Fso.CreateFile FilePath, o_value
        If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
      Case Else
        Easp.Error.FunctionName = "Cache:Item.Save"
        Easp.Error.Detail = Easp.Str.HtmlEncode(Me.Name)
        Easp.Error.Raise "error-cache-invalid-file"
    End Select
  End Sub
  '删除文件缓存
  Public Sub Remove
    '删除文件缓存
    If Not Easp.Str.Test(DelPath,"[*?]") Then
      If Easp.Fso.IsExists(DelPath) Then Easp.Fso.Del DelPath
      If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
    Else
      '如果有通配符
      Easp.Fso.DelFile left(DelPath,len(DelPath)-Len(FileType))
      Easp.Fso.DelFolder left(DelPath,len(DelPath)-Len(FileType))
      If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
    End If
  End Sub
  '删除内存缓存
  Public Sub RemoveApp
    If Easp.Has(Me.Name) Then Easp.RemoveApplication Me.Name
    If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
  End Sub
  '取文件缓存的缓存路径
  Public Property Get FilePath()
    FilePath = TransPath("[\\:""*?<>|\f\n\r\t\v\s]")
  End Property
  '取文件缓存的缓存地址，可带通配符
  Private Function DelPath()
    DelPath = TransPath("[\\:""<>|\f\n\r\t\v\s]")
  End Function
  '将名称转换为文件缓存地址
  Private Function TransPath(ByVal fe)
    Dim s_p : s_p = ""
    Dim parr : parr = split(Me.Name,"/")
    Dim i
    for i = 0 to UBound(parr)
      If Easp.Str.Test(parr(i),fe) Then parr(i)=Server.URLEncode(parr(i))
      s_p = s_p & "_" & parr(i)
      If i < UBound(parr) Then
        s_p = s_p & "\"
      End If
    next
    If s_p="" Then s_p="_"
    TransPath = SavePath & s_p & FileType
  End Function  
  '缓存是否可用（未过期）
  Public Function Ready()
    Dim app : app = Easp.GetApplication(Me.Name)
    Ready = False
    '如果是内存缓存
    If IsArray(app) Then
      If UBound(app) = 1 Then
        If IsDate(app(0)) Then
          Ready = isValid(app(0))
          If Ready Then Exit Function
        End If
      End If
    '如果是文件缓存
    ElseIf Easp.Fso.IsFile(FilePath) Then
      Ready = isValid(Easp.Fso.GetAttr(FilePath,1))
    End If
  End Function
  '验证时间是否过期
  Private Function isValid(ByVal t)
    If IsDate(t) Then
      If Easp.Has(d_exp) Then
        isValid = (DateDiff("s",Now,d_exp) > 0)
      Else
        isValid = (DateDiff("s",t,Now) < i_exp*60)
      End If
    End If
  End Function
End Class
%>