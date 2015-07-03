<%
'######################################################################
'## easp.stringobject.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP String Object Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-16 23:53:45
'## Description :   Format a string with chaining operations.
'##
'######################################################################
'链式操作Str方法
Class EasyASP_StringObject
  Private s_source
  '设置源
  Public Property Let Value(ByRef string)
    If IsObject(string) Then
      Set s_source = string
    Else
      s_source = string
    End If
  End Property
  '读取处理后的源
  Public Default Property Get Value()
    If IsObject(s_source) Then
      Set Value = s_source
    Else
      Value = s_source
    End If
  End Property
  Private Function S(ByRef string)
    Set S = New EasyASP_StringObject
    S.Value = string
  End Function

  Public Function Format(ByVal value)
    Set Format = S(Easp.Str.Format(s_source, value))
  End Function
  Public Function IsSame(ByVal string)
    IsSame = Easp.Str.IsSame(s_source, string)
  End Function
  Public Function IsEqual(ByVal string)
    IsEqual = Easp.Str.IsEqual(s_source, string)
  End Function
  Public Function Compare(ByVal t, ByVal b)
    Compare = Easp.Str.Compare(s_source, t, b)
  End Function
  Public Function IsIn(string)
    IsIn = Easp.Str.IsIn(s_source, string)
  End Function
  Public Function IsInList(ByVal string)
    IsInList = Easp.Str.IsInList(s_source, string)
  End Function
  Public Function StartsWith(ByVal string)
    StartsWith = Easp.Str.StartsWith(s_source, string)
  End Function
  Public Function EndsWith(ByVal string)
    EndsWith = Easp.Str.EndsWith(s_source, string)
  End Function
  Public Function GetColonName()
    Set GetColonName = S(Easp.Str.GetColonName(s_source))
  End Function
  Public Function GetColonValue()
    Set GetColonValue = S(Easp.Str.GetColonValue(s_source))
  End Function
  Public Function GetName(ByVal separator)
    Set GetName = S(Easp.Str.GetName(s_source, separator))
  End Function
  Public Function GetValue(ByVal separator)
    Set GetValue = S(Easp.Str.GetValue(s_source, separator))
  End Function
  Public Function GetNameValue(ByVal separator)
    Set GetNameValue = S(Easp.Str.GetNameValue(s_source, separator))
  End Function
  Public Function Cut(ByVal strlen)
    Set Cut = S(Easp.Str.Cut(s_source, strlen))
  End Function
  Public Function Replace(ByVal rule, ByVal replaceWith)
    Set Replace = S(Easp.Str.Replace(s_source, rule, replaceWith))
  End Function
  Public Function ReplaceLine(ByVal rule, ByVal replaceWith)
    Set ReplaceLine = S(Easp.Str.ReplaceLine(s_source, rule, replaceWith))
  End Function
  Public Function ReplacePart(ByVal rule, ByVal group, ByVal replaceWith)
    Set ReplacePart = S(Easp.Str.ReplacePart(s_source, rule, group, replaceWith))
  End Function
  Public Function Match(ByRef rule)
    Set Match = Easp.Str.Match(s_source, rule)
  End Function
  Public Function [Test](ByRef rule)
    [Test] = Easp.Str.Test(s_source, rule)
  End Function
  Public Function RegexpEncode()
    Set RegexpEncode = S(Easp.Str.RegexpEncode(s_source))
  End Function
  Public Function TrimChar(ByVal char)
    Set TrimChar = S(Easp.Str.TrimChar(s_source, char))
  End Function
  Public Function HtmlEncode()
    Set HtmlEncode = S(Easp.Str.HtmlEncode(s_source))
  End Function
  Public Function HtmlDecode()
    Set HtmlDecode = S(Easp.Str.HtmlDecode(s_source))
  End Function
  Public Function HtmlFilter()
    Set HtmlFilter = S(Easp.Str.HtmlFilter(s_source))
  End Function
  Public Function HtmlFormat()
    Set HtmlFormat = S(Easp.Str.HtmlFormat(s_source))
  End Function
  Public Function HtmlSafe()
    Set HtmlSafe = S(Easp.Str.HtmlSafe(s_source))
  End Function
  Public Function ToString()
    Set ToString = S(Easp.Str.ToString(s_source))
  End Function
  Public Function JsEncode()
    Set JsEncode = S(Easp.Str.JsEncode(s_source))
  End Function
  Public Function JsEncode_(ByVal cn)
    Set JsEncode_ = S(Easp.Str.JsEncode_(s_source, cn))
  End Function
  Public Function JavaScript()
    Set JavaScript = S(Easp.Str.JavaScript(s_source))
  End Function
  Public Sub JsAlert()
    Call Easp.Str.JsAlert(s_source)
  End Sub
  Public Sub JsAlertUrl(ByVal url)
    Call Easp.Str.JsAlertUrl(s_source, url)
  End Sub
  Public Sub JsConfirmUrl(ByVal yesUrl, ByVal cancelUrl)
    Call Easp.Str.JsConfirmUrl(s_source, yesUrl, cancelUrl)
  End Sub
  Public Function RandomStr()
    Set RandomStr = S(Easp.Str.RandomStr(s_source))
  End Function
  Public Function RandomString(ByVal allowStr)
    Set RandomString = S(Easp.Str.RandomString(s_source, allowStr))
  End Function
  Public Function RandomNumber(ByVal max)
    Set RandomNumber = S(Easp.Str.RandomNumber(s_source, max))
  End Function
  Public Function ToNumber(ByVal decimalType)
    Set ToNumber = S(Easp.Str.ToNumber(s_source, decimalType))
  End Function
  Public Function ToPrice()
    Set ToPrice = S(Easp.Str.ToPrice(s_source))
  End Function
  Public Function ToPercent()
    Set ToPercent = S(Easp.Str.ToPercent(s_source))
  End Function
  Public Function Half2Full()
    Set Half2Full = S(Easp.Str.Half2Full(s_source))
  End Function
  Public Function Full2Half()
    Set Full2Half = S(Easp.Str.Full2Half(s_source))
  End Function
  Public Function Validate()
    Set Validate = S(Easp.Str.Validate(s_source))
  End Function

  '将ASP函数重写为链式操作
  'Replace
  Public Function Rep(ByVal find, ByVal replacewith)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set Rep = S(o_re.Re(s_source, find, replaceWith))
    Set o_re = Nothing
  End Function
  'Replace 忽略大小写
  Public Function iReplace(ByVal find, ByVal replaceWith)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set iReplace = S(o_re.ReCase(s_source, find, replaceWith))
    Set o_re = Nothing
  End Function
  'Replace 完整参数
  Public Function RepAll(ByVal find, ByVal replaceWith, ByVal start, ByVal count, ByVal compare)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set RepAll = S(o_re.ReFull(s_source, find, replaceWith, start, count, compare))
    Set o_re = Nothing
  End Function
  Public Function Instr(ByVal string)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Instr = o_re.Instr_(s_source, string)
    Set o_re = Nothing
  End Function
  Public Function InstrAll(ByVal string, ByVal start, ByVal compare)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    InstrAll = o_re.Instr__(s_source, string, start, compare)
    Set o_re = Nothing
  End Function
  Public Function InStrRev(ByVal string)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    InStrRev = o_re.InStrRev_(s_source, string)
    Set o_re = Nothing
  End Function
  Public Function InStrRevAll(ByVal string, ByVal start, ByVal compare)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    InStrRevAll = o_re.InStrRev__(s_source, string, start, compare)
    Set o_re = Nothing
  End Function
  Public Function LCase()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set LCase = S(o_re.LCase_(s_source))
    Set o_re = Nothing
  End Function
  Public Function Left(ByVal length)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set Left = S(o_re.Left_(s_source, length))
    Set o_re = Nothing
  End Function
  Public Function Len()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Len = o_re.Len_(s_source)
    Set o_re = Nothing
  End Function
  Public Function LTrim()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set LTrim = S(o_re.LTrim_(s_source))
    Set o_re = Nothing
  End Function
  Public Function RTrim()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set RTrim = S(o_re.RTrim_(s_source))
    Set o_re = Nothing
  End Function
  Public Function Trim()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set Trim = S(o_re.Trim_(s_source))
    Set o_re = Nothing
  End Function
  Public Function Mid(ByVal start)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set Mid = S(o_re.Mid_(s_source, start))
    Set o_re = Nothing
  End Function
  Public Function MidAll(ByVal start, ByVal length)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set MidAll = S(o_re.Mid__(s_source, start, length))
    Set o_re = Nothing
  End Function
  Public Function Right(ByVal length)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set Right = S(o_re.Right_(s_source, length))
    Set o_re = Nothing
  End Function
  Public Function StrComp(ByVal string)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    StrComp = o_re.StrComp_(s_source, string)
    Set o_re = Nothing
  End Function
  Public Function StrCompAll(ByVal string, ByVal compare)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    StrCompAll = o_re.StrComp__(s_source, string, compare)
    Set o_re = Nothing
  End Function
  Public Function StrReverse()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set StrReverse = S(o_re.StrReverse_(s_source))
    Set o_re = Nothing
  End Function
  Public Function Split(ByVal separator)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Split = o_re.Split_(s_source, separator)
    Set o_re = Nothing
  End Function
  Public Function UCase()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Set UCase = S(o_re.UCase_(s_source))
    Set o_re = Nothing
  End Function

  Public Function CDate()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CDate = o_re.CDate_(s_source)
    Set o_re = Nothing
  End Function
  Public Function IsDate()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    IsDate = o_re.IsDate_(s_source)
    Set o_re = Nothing
  End Function
  Public Function Asc()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Asc = o_re.Asc_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CBool()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CBool = o_re.CBool_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CByte()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CByte = o_re.CByte_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CCur()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CCur = o_re.CCur_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CDbl()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CDbl = o_re.CDbl_(s_source)
    Set o_re = Nothing
  End Function
  Public Function Chr()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Chr = o_re.Chr_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CInt()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CInt = o_re.CInt_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CLng()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CLng = o_re.CLng_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CSng()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CSng = o_re.CSng_(s_source)
    Set o_re = Nothing
  End Function
  Public Function CStr()
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    CStr = o_re.CStr_(s_source)
    Set o_re = Nothing
  End Function
    
  Public Function Round(ByVal numdecimalplaces)
    Dim o_re : Set o_re = New EasyASP_StringOriginal
    Round = o_re.Round_(s_source, numdecimalplaces)
    Set o_re = Nothing
  End Function
End Class
%>