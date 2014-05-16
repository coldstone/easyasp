<%
'######################################################################
'## easp.validation.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP String Validation Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-05-12 0:38:21
'## Description :   With defined rules or custom rules to verify the 
'##                 legitimacy of a string.
'##
'######################################################################

Class EasyASP_Validation

  Private b_validate, b_return, s_source, s_value
  Private s_field, s_msg, s_msgDefault, s_default, s_name
  
  Private Sub Class_Initialize()
    b_validate   = True
    b_return     = True
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
      Value = Easp.IIF(b_validate, Easp.IfHas(s_source, s_default), Easp.IfHas(s_default, Easp.IfHas(s_msg, False)))
    End If
  End Property
  '生成验证对象
  Private Function NewValidation()
    Set NewValidation = New EasyASP_Validation
    NewValidation.Source = s_source
    NewValidation.NameString = s_name
    NewValidation.ReturnValue = b_return
    NewValidation.PostField = s_field
    NewValidation.Validate = b_validate
    NewValidation.MsgInfo = s_msg
    NewValidation.DefaultValue = s_default
    NewValidation.MsgDefault = s_msgDefault
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
    Set Required = NewValidation()
    If Easp.IsN(s_source) Then
      Required.Validate = False
      CreateMsgDefault Required, "required", Null
    End If
  End Function
  '验证正则规则
  Public Function Test(ByVal regexpString)
    Set Test = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, regexpString) Then
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
    Set DateBetween = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "date") Then
      DateBetween.Validate = False
      CreateMsgDefault DateBetween, "isdate", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDate(s_source) < CDate(minDate) Or CDate(s_source) > CDate(maxDate) Then
        DateBetween.Validate = False
        CreateMsgDefault DateBetween, "datebetween", Array(minDate, maxDate)
      End If
    End If
  End Function
  '验证最小日期
  Public Function MinDate(ByVal dateTime)
    Set MinDate = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "date") Then
      MinDate.Validate = False
      CreateMsgDefault MinDate, "isdate", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDate(s_source) < CDate(dateTime) Then
        MinDate.Validate = False
        CreateMsgDefault MinDate, "mindate", dateTime
      End If
    End If
  End Function
  '验证最大日期
  Public Function MaxDate(ByVal dateTime)
    Set MaxDate = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "date") Then
      MaxDate.Validate = False
      CreateMsgDefault MaxDate, "isdate", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDate(s_source) > CDate(dateTime) Then
        MaxDate.Validate = False
        CreateMsgDefault MaxDate, "maxdate", dateTime
      End If
    End If
  End Function
  '验证数值
  Public Function IsNumber()
    Set IsNumber = Me.Test("number")
  End Function
  '验证数值区间
  Public Function Between(ByVal min, ByVal max)
    Set Between = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "number") Then
      Between.Validate = False
      CreateMsgDefault Between, "isnumber", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDbl(s_source) < CDbl(Min) Or CDbl(s_source) > CDbl(Max) Then
        Between.Validate = False
        CreateMsgDefault Between, "between", Array(min, max)
      End If
    End If
  End Function
  '验证最小数值
  Public Function Min(ByVal number)
    Set Min = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "number") Then
      Min.Validate = False
      CreateMsgDefault Min, "isnumber", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDbl(s_source) < CDbl(number) Then
        Min.Validate = False
        CreateMsgDefault Min, "min", number
      End If
    End If
  End Function
  '验证最大数值
  Public Function Max(ByVal number)
    Set Max = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.Test(s_source, "number") Then
      Max.Validate = False
      CreateMsgDefault Max, "isnumber", Null
      Exit Function
    End If
    If Easp.Has(s_source) Then
      If CDbl(s_source) > CDbl(number) Then
        Max.Validate = False
        CreateMsgDefault Max, "max", number
      End If
    End If
  End Function
  '验证长度区间
  Public Function LengthBetween(ByVal minNumber, ByVal maxNumber)
    Set LengthBetween = NewValidation()
    If Easp.Has(s_source) And Len(s_source) < minNumber Or Len(s_source) > maxNumber Then
      LengthBetween.Validate = False
      CreateMsgDefault LengthBetween, "lengthbetween", Array(minNumber, maxNumber)
    End If
  End Function
  '验证长度
  Public Function Length(ByVal number)
    Set Length = NewValidation()
    If Easp.Has(s_source) And Len(s_source) <> number Then
      Length.Validate = False
      CreateMsgDefault Length, "length", number
    End If
  End Function
  '验证最小长度
  Public Function MinLength(ByVal number)
    Set MinLength = NewValidation()
    If Easp.Has(s_source) And Len(s_source) < number Then
      MinLength.Validate = False
      CreateMsgDefault MinLength, "minlength", number
    End If
  End Function
  '验证最大长度
  Public Function MaxLength(ByVal number)
    Set MaxLength = NewValidation()
    If Easp.Has(s_source) And Len(s_source) > number Then
      MaxLength.Validate = False
      CreateMsgDefault MaxLength, "maxlength", number
    End If
  End Function
  '验证相等
  Public Function Same(ByVal string)
    Set Same = NewValidation()
    If Easp.Has(s_source) And Not Easp.Str.IsSame(s_source, string) Then
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
%>