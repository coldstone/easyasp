<%
'######################################################################
'## easp.var.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Variables Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-18 11:15:17
'## Description :   Get and set EasyASP super variables.
'##
'######################################################################

Class EasyASP_Var
  Private o_var
  Private b_loaded
  Private Sub Class_Initialize()
    Set o_var = Server.CreateObject("Scripting.Dictionary")
    o_var.CompareMode = 1 '对key进行文本比较
    b_loaded = False
  End Sub
  Private Sub Class_Terminate()
    Dim key
    For Each key In o_var
      If IsObject(o_var(key)) Then Set o_var(key) = Nothing
    Next
    Set o_var = Nothing
  End Sub
  
  '读取和设置EasyASP超级变量集
  '优先级依次为: 
  '  1.自定义变量
  '  2.Request.QueryString
  '  3.Request.Form
  '  4.EasyASP系统变量
  '  5.Request.ServerVariables
  '  即如有同名的变量，在上一个集合中找到则立即返回值，不再在下一集合中查找
  Public Default Property Get Var(ByVal key)
    Dim s_var
    If Not b_loaded Then Call getVars()
    If o_var.Exists(key) Then
      s_var = o_var(key)
    ElseIf o_var.Exists("get." & key) Then
      s_var = o_var("get." & key)
    ElseIf o_var.Exists("post." & key) Then
      s_var = o_var("post." & key)
    ElseIf Easp.Str.IsSame(key,"easp.newid") Then
      s_var = Easp.NewID()
    ElseIf Easp.Str.IsSame(key,"easp.datetime") Then
      s_var = Now()
    ElseIf Easp.Str.IsSame(key,"easp.date") Then
      s_var = Date()
    ElseIf Easp.Str.IsSame(key,"easp.time") Then
      s_var = Time()
    ElseIf o_var.Exists("easp." & key) Then
      s_var = o_var("easp." & key)
    ElseIf Easp.Str.StartsWith(key, "server.") Then
      s_var = Request.ServerVariables(Mid(key,8))
    Else
      s_var = ""
    End If
    If Not IsArray(s_var) Then
      If Instr(s_var, "{=") Then
        Dim matches, match
        Set matches = Easp.Str.Match(s_var, "\{=(.+?)\}")
        For Each match In matches
          s_var = Replace(s_var, match, Me.Var(match.SubMatches(0)), 1, -1, 1)
        Next
        Set matches = Nothing
      End If
    End If
    Var = s_var
  End Property
  Public Property Let Var(ByVal key, ByVal value)
    If Not b_loaded Then Call getVars()
    Dim i,leng
    If IsArray(value) Then
      leng = UBound(value)
      '如果是数组则：1、将数组转为字符串存入 key
      o_var(key) = Join(value, ", ")
      '2、将原始数组存入 key_array
      o_var(key & "_array") = value
      '3、将数组长度存入 key_array_length
      o_var(key & "_array_length") = leng + 1
      '4、将数组元素存入 key_array_0、key_array_1...
      For i = 0 To leng
        o_var(key & "_array_" & i) = value(i)
      Next
    Else
      o_var(key) = value
    End If
  End Property

  '取得EasyASP变量集原始字典对象
  Public Function [GetObject]()
    If Not b_loaded Then Call getVars()
    Set [GetObject] = o_var
  End Function

  '查找是否包含某一变量
  Public Function Has(ByVal key)
    If Not b_loaded Then Call getVars()
    Has = (o_var.Exists(key) Or o_var.Exists("get." & key) Or o_var.Exists("post." & key) Or Easp.Str.IsInList("easp.newid,easp.datetime,easp.date,easp.time", key) Or o_var.Exists("easp." & key))
  End Function
  
  '将页面参数获取值写入到Var集合
  Private Sub getVars()
    ''取表单值
    If Not Easp.Upload.checkEntryType Then
      Call GetRequest("post")
    Else
      If Not Easp.Upload.IsUploaded Then Easp.Upload.GetData()
      Dim o_dic, s_post, key, value
      Set o_dic = Easp.Upload.Post("-1")
      If o_dic.Count > 0 Then
        For Each s_post In o_dic
          key = "post." & s_post
          value = o_dic(s_post)
          o_var(key) = value
        Next
      End If
      'Easp.Console o_var
    End If
    ''取URL参数值
    Call GetRequest("get")
    b_loaded = True
  End Sub
  '取得Post和Get值
  Private Sub GetRequest(ByVal requestType)
    Dim requestDic, requestKey
    Dim key, value, valueCount, values(), i
    If requestType = "get" Then
      Set requestDic = Request.QueryString
    ElseIf requestType = "post" Then
      Set requestDic = Request.Form
    End If
    If requestDic.Count > 0 Then
      For Each requestKey In requestDic
        key = requestType & "." & LCase(requestKey)
        value =  requestDic(requestKey)
        o_var(key) = value
        '同时存入数组***_array变量
        valueCount = requestDic(requestKey).Count - 1
        key = key & "_array"
        o_var(key & "_length") = valueCount + 1
        ReDim values(valueCount)
        For i = 0 To (valueCount)
          values(i) = requestDic(requestKey)(i+1)
          o_var(key & "_" & i) = values(i)
        Next
        o_var(key) = values
      Next
    End If
  End Sub
End Class
%>