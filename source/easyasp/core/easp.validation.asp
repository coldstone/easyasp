<%
'######################################################################
'## easp.validation.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP String Validation Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-22 0:48:20
'## Description :   With defined rules or custom rules to verify the 
'##                 legitimacy of a string.
'##
'######################################################################

Class EasyASP_Validation

  Private b_validate, b_return, s_source, s_split, a_return, b_array
  Private s_value, s_field, s_msg, s_msgDefault, s_default, s_name
  
  Private Sub Class_Initialize()
    b_validate   = True
    b_return     = True
    b_array      = False
    s_split      = Empty
    s_name       = Empty
    s_value      = Empty
    s_msgDefault = Empty
    s_field      = Empty
    s_msg        = Empty
    s_default    = Empty
  End Sub
  '设置源验证文本
  Public Property Let Source(ByVal string)
    s_source = string
    If Not IsArray(a_return) Then a_return = Array(s_source)
  End Property
  Public Property Get Source
    Source = s_source
  End Property
  '设置和获取验证结果
  Public Property Get Validate()
    Validate = b_validate
  End Property
  Public Property Let Validate(ByVal bool)
    b_validate = bool
  End Property
  '设置和获取错误提示
  Public Property Get MsgInfo()
    MsgInfo = s_msg
  End Property
  Public Property Let MsgInfo(ByVal string)
    s_msg = string
  End Property
  '设置和获取规则自带默认错误提示
  Public Property Get MsgDefault()
    MsgDefault = s_msgDefault
  End Property
  Public Property Let MsgDefault(ByVal string)
    s_msgDefault = string
  End Property
  '设置和获取是否有返回值
  Public Property Get ReturnValue()
    ReturnValue = b_return
  End Property
  Public Property Let ReturnValue(ByVal bool)
    b_return = bool
  End Property
  '设置和获取分隔符
  Public Property Get Separator()
    Separator = s_split
  End Property
  Public Property Let Separator(ByVal string)
    s_split = string
  End Property
  '设置和获取是否返回数组
  Public Property Get IsReturnArray()
    IsReturnArray = b_array
  End Property
  Public Property Let IsReturnArray(ByVal bool)
    b_array = bool
  End Property
  '设置和获取返回数组
  Public Property Get ReturnArray()
    ReturnArray = a_return
  End Property
  Public Property Let ReturnArray(ByVal arr)
    a_return = arr
  End Property
  '设置和获取名称
  Public Property Get NameString()
    NameString = s_name
  End Property
  Public Property Let NameString(ByVal string)
    s_name = string
  End Property
  '设置和获取表单名
  Public Property Get PostField()
    PostField = s_field
  End Property
  Public Property Let PostField(ByVal string)
    s_field = string
  End Property
  '设置和获取默认值
  Public Property Get DefaultValue()
    DefaultValue = s_default
  End Property
  Public Property Let DefaultValue(ByVal string)
    s_default = string
  End Property
  '获取返回值
  Public Default Property Get Value()
    If b_return Then
      Dim i
      For i = 0 To UBound(a_return)
        a_return(i) = CStr(Easp.IIF(b_validate, Easp.IfHas(a_return(i), s_default), Easp.IfHas(a_return(i), Easp.IfHas(s_default, Easp.IfHas(s_msg, False)))))
      Next
      If b_array Then
        Value = a_return
      Else
        Dim o : Set o = New EasyASP_ValidationTool
        Value = o.Join_(a_return, s_split)
        Set o = Nothing
      End If
    End If
  End Property
  '生成验证对象
  Private Function NewValidation()
    Set NewValidation = New EasyASP_Validation
    NewValidation.Source = s_source '继承源文本
    NewValidation.NameString = s_name '继承名称
    NewValidation.ReturnValue = b_return '继承返回值
    NewValidation.PostField = s_field '继承表单项名称
    NewValidation.Validate = b_validate '继承验证是否通过
    NewValidation.MsgInfo = s_msg '继承提示信息
    NewValidation.DefaultValue = s_default '继承默认值
    NewValidation.MsgDefault = s_msgDefault '继承默认提示信息
    NewValidation.Separator = s_split '继承分隔符
    NewValidation.IsReturnArray = b_array '继承是否返回数组
    NewValidation.ReturnArray = a_return '继承返回的数组
  End Function
  '去除两头空白
  Public Function Trim()
    Dim o : Set o = New EasyASP_StringOriginal
    s_source = o.Trim_(s_source)
    Set o = Nothing
    Set Trim = NewValidation()
  End Function
  '设置默认值
  Public Function [Default](ByVal string)
    Set [Default] = NewValidation()
    [Default].DefaultValue = string
  End Function
  '设置名称
  Public Function [Name](ByVal string)
    Set [Name] = NewValidation()
    [Name].NameString = string
  End Function
  '设置分隔符(设置后会按分隔符分隔后一项项验证)
  Public Function Split(ByVal string)
    Dim o : Set o = New EasyASP_StringOriginal
    a_return = o.Split_(s_source, string)
    s_split = string
    Set o = Nothing
    Set Split = NewValidation()
  End Function
  '将分隔后的数组按另外的符号组合为字符串
  Public Function Join(ByVal string)
    s_split = string
    Set Join = NewValidation()
  End Function
  '设置返回数组
  Public Function GetArray()
    b_array = True
    Set GetArray = NewValidation()
  End Function
  '设置不返回数据
  Public Function NoReturn()
    Set NoReturn = NewValidation()
    NoReturn.ReturnValue = False
  End Function
  '设置表单名
  Public Function [Field](ByVal string)
    Set [Field] = NewValidation()
    [Field].PostField = string
  End Function
  '设置上一规则错误提示信息
  Public Function Msg(ByVal string)
    Set Msg = NewValidation()
    If b_validate Then
      Msg.MsgInfo = s_msg
    Else
      If Easp.Has(s_msg) Then
        Msg.MsgInfo = s_msg
      Else
        If Instr(string, "%n") Then
         string = Replace(string, "%n", Easp.IfHas(s_name, Easp.Lang("val-item")))
        End If
        If Instr(string, "%s") Then string = Replace(string, "%s", s_source)
        Msg.MsgInfo = string
      End If
    End If
  End Function
  '规则验证失败则弹出javascript警告框
  Public Function Alert()
    Set Alert = NewValidation()
    If Not b_validate Then
      Easp.Str.JsAlert Easp.IfHas(s_msg, s_msgDefault)
    End If
  End Function
  '规则验证失败则弹出javascript警告框并跳转到新页面
  Public Function AlertUrl(ByVal url)
    Set AlertUrl = NewValidation()
    If Not b_validate Then
      Easp.Str.JsAlertUrl Easp.IfHas(s_msg, s_msgDefault), url
    End If
  End Function
  '规则验证失败则打印出错误提示信息并终止程序运行
  Public Function PrintEnd()
    Set PrintEnd = NewValidation()
    If Not b_validate Then
      Easp.PrintEnd Easp.IfHas(s_msg, s_msgDefault)
    End If
  End Function
  '规则验证失败则打印出Json格式错误提示信息并终止程序运行
  Public Function PrintEndJson()
    Set PrintEndJson = NewValidation()
    If Not b_validate Then
      Dim s_json
      s_json = "{""validateError"" : """
      s_json = s_json & Easp.Str.JsEncode(Easp.IfHas(s_msg, s_msgDefault))
      If Easp.Has(s_field) Then
        s_json = s_json & """, ""validateField"" : """
        s_json = s_json & s_field
      End If
      s_json = s_json & """}"
      Easp.PrintEnd s_json
    End If
  End Function
  '生成验证规则默认提示信息
  Private Sub CreateMsgDefault(ByRef validation, ByVal langString, ByVal arrValue)
    Dim string, i
    If Easp.IsN(validation.MsgDefault) Then
      string = Easp.IfHas(validation.NameString, Easp.Lang("val-item"))
      If Easp.Lang.Exists("val-" & langString) Then
        '如果语言包中有相应的内容
        string = string & Easp.Lang("val-" & langString)
        If Easp.Has(arrValue) Then
          If Not IsArray(arrValue) Then arrValue = Array(arrValue)
          For i = 0 To UBound(arrValue)
            string = Replace(string, "%v", arrValue(i), 1, 1, 1)
          Next
        End If
      Else
        string = string & Easp.Lang("val-test")
      End If
      validation.MsgDefault = string
    End If
  End Sub
  '验证不能为空
  Public Function Required()
    Dim i, b_val : b_val = True
    If UBound(a_return) < 0 Then b_val = False
    For i = 0 To UBound(a_return)
      If Easp.IsN(a_return(i)) Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set Required = NewValidation()
    If Not b_val Then
      Required.Validate = False
      CreateMsgDefault Required, "required", Null
    End If
  End Function
  '验证正则规则
  Public Function Test(ByVal regexpString)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), regexpString) Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set Test = NewValidation()
    If Not b_val Then
      Test.Validate = False
      CreateMsgDefault Test, "test-" & regexpString, Null
    End If
  End Function
  '验证日期
  Public Function [IsDate]()
    Set [IsDate] = Me.Test("date")
  End Function
  '验证日期区间
  Public Function DateBetween(ByVal minDate, ByVal maxDate)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "date") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDate(a_return(i)) < CDate(minDate) Or CDate(a_return(i)) > CDate(maxDate) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set DateBetween = NewValidation()
    If Not b_val1 Then
      DateBetween.Validate = False
      CreateMsgDefault DateBetween, "isdate", Null
      Exit Function
    End If
    If Not b_val2 Then
      DateBetween.Validate = False
      CreateMsgDefault DateBetween, "datebetween", Array(minDate, maxDate)
    End If
  End Function
  '验证最小日期
  Public Function MinDate(ByVal dateTime)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "date") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDate(a_return(i)) < CDate(dateTime) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set MinDate = NewValidation()
    If Not b_val1 Then
      MinDate.Validate = False
      CreateMsgDefault MinDate, "isdate", Null
      Exit Function
    End If
    If Not b_val2 Then
      MinDate.Validate = False
      CreateMsgDefault MinDate, "mindate", dateTime
    End If
  End Function
  '验证最大日期
  Public Function MaxDate(ByVal dateTime)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "date") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDate(a_return(i)) > CDate(dateTime) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set MaxDate = NewValidation()
    If Not b_val1 Then
      MaxDate.Validate = False
      CreateMsgDefault MaxDate, "isdate", Null
      Exit Function
    End If
    If Not b_val2 Then
      MaxDate.Validate = False
      CreateMsgDefault MaxDate, "maxdate", dateTime
    End If
  End Function
  '验证数值
  Public Function IsNumber()
    Set IsNumber = Me.Test("number")
  End Function
  '验证数值区间
  Public Function Between(ByVal min, ByVal max)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "number") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDbl(a_return(i)) < CDbl(Min) Or CDbl(a_return(i)) > CDbl(Max) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set Between = NewValidation()
    If Not b_val1 Then
      Between.Validate = False
      CreateMsgDefault Between, "isnumber", Null
      Exit Function
    End If
    If Not b_val2 Then
      Between.Validate = False
      CreateMsgDefault Between, "between", Array(min, max)
    End If
  End Function
  '验证最小数值
  Public Function Min(ByVal number)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "number") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDbl(a_return(i)) < CDbl(number) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set Min = NewValidation()
    If Not b_val1 Then
      Min.Validate = False
      CreateMsgDefault Min, "isnumber", Null
      Exit Function
    End If
    If Not b_val2 Then
      Min.Validate = False
      CreateMsgDefault Min, "min", number
    End If
  End Function
  '验证最大数值
  Public Function Max(ByVal number)
    Dim i, b_val1, b_val2
    b_val1 = True
    b_val2 = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.Test(a_return(i), "number") Then
        a_return(i) = Empty
        If b_val1 Then b_val1 = False
      End If
      If Easp.Has(a_return(i)) Then
        If CDbl(a_return(i)) > CDbl(number) Then
          a_return(i) = Empty
          If b_val2 Then b_val2 = False
        End If
      End If
    Next
    Set Max = NewValidation()
    If Not b_val1 Then
      Max.Validate = False
      CreateMsgDefault Max, "isnumber", Null
      Exit Function
    End If
    If Not b_val2 Then
      Max.Validate = False
      CreateMsgDefault Max, "max", number
    End If
  End Function
  '验证长度区间
  Public Function LengthBetween(ByVal minNumber, ByVal maxNumber)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Len(a_return(i)) < minNumber Or Len(a_return(i)) > maxNumber Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set LengthBetween = NewValidation()
    If Not b_val Then
      LengthBetween.Validate = False
      CreateMsgDefault LengthBetween, "lengthbetween", Array(minNumber, maxNumber)
    End If
  End Function
  '验证长度
  Public Function Length(ByVal number)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Len(a_return(i)) <> number Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set Length = NewValidation()
    If Not b_val Then
      Length.Validate = False
      CreateMsgDefault Length, "length", number
    End If
  End Function
  '验证最小长度
  Public Function MinLength(ByVal number)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Len(a_return(i)) < number Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set MinLength = NewValidation()
    If Not b_val Then
      MinLength.Validate = False
      CreateMsgDefault MinLength, "minlength", number
    End If
  End Function
  '验证最大长度
  Public Function MaxLength(ByVal number)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Len(a_return(i)) > number Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set MaxLength = NewValidation()
    If Not b_val Then
      MaxLength.Validate = False
      CreateMsgDefault MaxLength, "maxlength", number
    End If
  End Function
  '验证相等
  Public Function Same(ByVal string)
    Dim i, b_val : b_val = True
    For i = 0 To UBound(a_return)
      If Easp.Has(a_return(i)) And Not Easp.Str.IsSame(a_return(i), string) Then
        a_return(i) = Empty
        If b_val Then b_val = False
      End If
    Next
    Set Same = NewValidation()
    If Not b_val Then
      Same.Validate = False
      CreateMsgDefault Same, "same", Null
    End If
  End Function
  '验证两次输入一致
  Public Function SamePost(ByVal string)
    Set SamePost = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.IsSame(s_source, Easp.Post(string)) Then
      SamePost.Validate = False
      CreateMsgDefault SamePost, "samepost", Null
    End If
  End Function
  '验证验证码输入
  Public Function SameSession(ByVal string)
    Set SameSession = NewValidation()
    If Not Easp.Str.IsSame(s_source, Session(string)) Or Easp.IsN(s_source) Then
      If Easp.IsN(s_name) Then
        SameSession.NameString = Easp.Lang("val-verify")
      End If
      SameSession.Validate = False
      CreateMsgDefault SameSession, "samesession", Null
    End If
  End Function
  
End Class
Class EasyASP_ValidationTool
  Public Function Join_(ByRef arr, ByVal separator)
    Join_ = Join(arr, separator)
  End Function
End Class
%>