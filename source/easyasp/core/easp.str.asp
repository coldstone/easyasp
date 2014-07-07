<%
'######################################################################
'## easp.str.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP String Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-13 17:00:00
'## Description :   EasyASP String Class
'##
'######################################################################

Class EasyASP_String
  Private o_re, b_encodeJson
  Private Sub Class_Initialize()
    Set o_re = New EasyASP_StringOriginal
  End Sub
  Private Sub Class_Terminate()
    Set o_re = Nothing
  End Sub

  '是否编码ToString时的Unicode字符
  Public Property Get EncodeJsonUnicode
    EncodeJsonUnicode = b_encodeJson
  End Property
  Public Property Let EncodeJsonUnicode(ByRef bool)
    b_encodeJson = bool
  End Property
  
  '格式化字符串（首下标为0）
  Public Function Format(ByVal string, ByVal value)
    Format = FormatString(string, value, 0)
  End Function
  Private Function FormatString(ByVal s, ByRef v, ByVal t)
    Dim i,n,k
    s = o_re.Re(s,"\\",Chr(0))
    s = o_re.Re(s,"\{",Chr(1))
    Select Case VarType(v)
      '数组
      Case 8192,8194,8204,8209
        For i = 0 To Ubound(v)
          s = FormatReplace(s,i+t,v(i))
        Next
      '对象
      Case 9
        Select Case TypeName(v)
          '记录集
          Case "Recordset"
            For i = 0 To v.Fields.Count - 1
              s = FormatReplace(s,i+t,v(i))
              s = FormatReplace(s,v.Fields.Item(i+t).Name,v(i))
            Next
          '字典
          Case "Dictionary", "EasyASP_Json_Object"
            If TypeName(v) = "EasyASP_Json_Object" Then Set v = v.GetDictionary
            For Each k In v
              s = FormatReplace(s,k,v(k))
            Next
          'Easp List
          Case "EasyASP_List"
            For i = 0 To v.End
              s = FormatReplace(s,i+t,v(i))
              s = FormatReplace(s,v.IndexHash(i),v(i))
            Next
          '正则搜索子集合
          Case "ISubMatches", "SubMatches"
            For i = 0 To v.Count - 1
              s = FormatReplace(s,i+t,v(i))
            Next
          Case "EasyASP_Json_Array"
            For i = 0 To v.Length - 1
              s = FormatReplace(s, i+t, v(i))
            Next
        End Select
      '字符串
      Case 8
        Select Case TypeName(v)
          '正则搜索集合
          Case "IMatch2", "Match"
            s = FormatReplace(s,t,v.Value)
            For i = 0 To v.SubMatches.Count - 1
              s = FormatReplace(s,i+t+1,v.SubMatches(i))
            Next
          '字符串
          Case Else
            s = FormatReplace(s,t,v)
        End Select
      Case Else
        s = FormatReplace(s,t,v)
    End Select
    If Instr(s, "{=") Then
      Dim matches, m
      Set matches = Match(s, "\{=(.+?)\}")
      For Each m In matches
        s = o_re.ReCase(s, m, Easp.Var(m.SubMatches(0)))
      Next
      Set matches = Nothing
    End If
    s = o_re.Re(s,Chr(1),"{")
    FormatString = o_re.Re(s,Chr(0),"\")
  End Function
  '格式化Format内标签参数
  Private Function FormatReplace(ByVal s, ByVal t, ByVal v)
    Dim tmp,rule,ru,kind,matches,ma
    v = Easp.IfHas(v,"")
    rule = "\{" & t & "(:((N[,\(%]?(\d+)?)|(D[^\}]+)|(E[^\}]+)|U|L|\d+([^\}]+)?))\}"
    If Me.Test(s,rule) Then
      Set matches = Me.Match(s,rule)
      For Each ma In matches
        kind = Replace(ma.Value, rule, "$2")
        ru = ma.Value
        Select Case Left(kind,1)
          '截取字符串
          Case "1","2","3","4","5","6","7","8","9"
            s = o_re.Re(s, ru, Cut(v,Replace(kind,"^(\d+)(.+)?$","$1:$2")))
          '数字
          Case "N"
            If isNumeric(v) Then
              Dim style,group,parens,percent,deci
              style = Replace(kind,"^N([,\(%])?(\d+)?$","$1")
              If style = "," Then group = -1
              If style = "(" Then parens = -1
              If style = "%" Then percent = -1
              deci = Replace(kind,"^N([,\(%])?(\d+)?$","$2")
              If Easp.IsN(style) And Easp.IsN(deci) Then
                s = o_re.ReCase(s, ru, Easp.IIF(Instr(Cstr(v),".")>0 And v<1,"0" & v,v))
              Else
                deci = Easp.IfHas(deci,-1)
                If percent Then
                  s = o_re.ReCase(s, ru, FormatNumber(v*100,deci,-1) & "%")
                Else
                  s = o_re.ReCase(s, ru, FormatNumber(v,deci,-1,parens,group))
                End If
              End If
            End If
          '日期
          Case "D"
            If isDate(v) Then
              s = o_re.ReCase(s, ru, Easp.Date.Format(v,Mid(kind,2)))
            End If
          '转大写
          Case "U"
            s = o_re.ReCase(s, ru, UCase(v))
          '转小写
          Case "L"
            s = o_re.ReCase(s, ru, LCase(v))
          '表达式
          Case "E"
            tmp = o_re.Re(Mid(kind,2), "%s", "v")
            tmp = Eval(tmp)
            s = o_re.ReCase(s, ru, tmp)
        End Select
      Next
    Else
      s = o_re.ReCase(s,"{" & t & "}",v)
    End If
    FormatReplace = s
  End Function
  
  '比较文本是否一致（忽略大小写）
  Public Function IsSame(ByVal string1, ByVal string2)
    string1 = Easp.IfHas(string1, "")
    string2 = Easp.IfHas(string2, "")
    IsSame = (StrComp(string1, string2, 1) = 0)
  End Function
  '比较文本是否一致（区分大小写）
  Public Function IsEqual(ByVal string1, ByVal string2)
    string1 = Easp.IfHas(string1, "")
    string2 = Easp.IfHas(string2, "")
    IsEqual = (StrComp(string1, string2, 0) = 0)
  End Function

  '比较两个字符串的大小，区分大小写
  '返回： Boolean值
  '用法： Easp.Print Compare("ABCD", "<", "abcd")  'True
  Public Function Compare(ByVal a, ByVal t, ByVal b)
    Dim isStr, b_comp
    isStr = False
    If VarType(a) = 8 Or VarType(b) = 8 Then
      isStr = True
      If IsNumeric(a) And IsNumeric(b) Then isStr = False
      If IsDate(a) And IsDate(b) Then isStr = False
    End If
    If isStr Then
      b_comp = StrComp(a,b,0)
      Select Case LCase(t)
        Case "lt", "<" Compare = (b_comp = -1)
        Case "gt", ">" Compare = (b_comp = 1)
        Case "eq", "=" Compare = (b_comp = 0)
        Case "lte", "<=" Compare = (b_comp = -1 Or b_comp = 0)
        Case "gte", ">=" Compare = (b_comp = 1 Or b_comp = 0)
      End Select
    Else
      Select Case LCase(t)
        Case "lt", "<" Compare = (a < b)
        Case "gt", ">" Compare = (a > b)
        Case "eq", "=" Compare = (a = b)
        Case "lte", "<=" Compare = (a <= b)
        Case "gte", ">=" Compare = (a >= b)
      End Select
    End If
  End Function

  '判断字符串中是否包含某字符串（忽略大小写）
  Public Function IsIn(string1, string2)
    If Easp.Has(string2) Then
      IsIn = InStr(1, string1, string2, 1)>0
    End If
  End Function

  '检查字符串是否属于逗号隔开的字符串序列中的一个
  '返回：Boolean值
  '说明：If Easp.Str.IsInList("A,B,C", str) Then
  '     等同于下面的语句并且忽略大小写：
  '     If str = "A" Or str = "B" Or str = "C" Then
  Public Function IsInList(ByVal string, ByVal str)
    Dim s1, s2
    If Easp.Has(str) Then
      s1 = Easp.IIF(Left(string,1)=",", string, "," & string)
      s1 = Easp.IIF(Right(s1,1)=",", s1, s1 & ",")
      s2 = Easp.IIF(Left(str,1)=",", str, "," & str)
      s2 = Easp.IIF(Right(s2,1)=",", s2, s2 & ",")
      IsInList = IsIn(s1, s2)
    End If
  End Function

  '检查字符串的开头是否与另一个字符串匹配
  Public Function StartsWith(ByVal string1, ByVal string2)
    'StartsWith = Test(string1, "^" & string2)
    StartsWith = IsSame(Left(string1, Len(string2)), string2)
  End Function
  '检查字符串的结尾是否与另一个字符串匹配
  Public Function EndsWith(ByVal string1, ByVal string2)
    'EndsWith = Test(string1, string2 & "$")
    EndsWith = IsSame(Right(string1, Len(string2)), string2)
  End Function

  '取“A:B”中的A
  Public Function GetColonName(ByVal string)
    GetColonName = GetNameValue(string, ":")(0)
  End Function
  '取“A:B”中的B
  Public Function GetColonValue(ByVal string)
    GetColonValue = GetNameValue(string, ":")(1)
  End Function
  '取“A分隔符B”中的A
  Public Function GetName(ByVal string, ByVal separator)
    GetName = GetNameValue(string, separator)(0)
  End Function
  '取“A分隔符B”中的B
  Public Function GetValue(ByVal string, ByVal separator)
    GetValue = GetNameValue(string, separator)(1)
  End Function
  '取分隔符字符串的两头
  '说明：把“A分隔符B”转为数组 Array(A,B)
  '返回：数组
  Public Function GetNameValue(ByVal string, ByVal separator)
    Dim n, arr(1)
    n = Instr(string, separator)
    If n > 0 Then
      arr(0) = Left(string, n-1)
      arr(1) = Mid(string, n+Len(separator))      
    Else
      arr(0) = string
      arr(1) = ""
    End If
    GetNameValue = arr
  End Function

  '截取长字符串左边部分并以特殊符号代替
  '半角字符以半个字符计，返回的字符串最大长度为strlen
  Public Function Cut(ByVal s, ByVal strlen)
    If Easp.IsN(s) Then Cut = "" : Exit Function
    '去除html标签、换行和制表符
    s = HtmlFilter(s)
    s = o_re.Re(s, vbCrLf, "")
    s = o_re.Re(s, vbTab, "")
    If Easp.IsN(strlen) Then Cut = s : Exit Function
    Dim i_len, i_leng, s_sep, i_sepleng, i_sleng, i, t
    '取出省略符
    If Instr(strlen,":") Then
      i_len = CInt(GetColonName(strlen))
      s_sep = GetColonValue(strlen)
      i_sepleng = Len(s_sep)
    Else
      i_len = strlen
      s_sep = "..."
      i_sepleng = 1
    End If
    '如果字符串不够长或者strlen为0，则返回全部字符串
    If i_len = 0 Or Len(s) <= i_len Then Cut = s : Exit Function
    '解决中英文混排长度
    i_sleng = Leng(s)
    i_leng = i_len * 2
    If i_sleng <= i_leng Then Cut = s : Exit Function
    i_leng = i_leng - i_sepleng * 2
    t = 0
    For i = 1 To Len(s)
      t = Easp.IIF(Abs(Ascw(Mid(s,i,1)))>255, t+2, t+1)
      If t >= i_leng Then
        Cut = Left(s, i) & s_sep
        Exit For
      End If
    Next
  End Function
  '返回字符串的长度，中文算两个字符
  Private Function Leng(string)
    Dim i,n : n = 0
    For i = 1 To Len(string)
      n = Easp.IIF(Abs(Ascw(Mid(string,i,1)))>255, n+2, n+1)
    Next
    Leng = n
  End Function

  '正则替换
  Public Function Replace(ByVal string, ByVal rule, ByVal replaceWith)
    Replace = Easp_Replace(string, rule, replaceWith, False)
  End Function
  '正则替换多行模式
  Public Function ReplaceLine(ByVal string, ByVal rule, ByVal replaceWith)
    ReplaceLine = Easp_Replace(string, rule, replaceWith, True)
  End Function
  '替换正则表达式编组
  '说明：按正则表达式的规则替换一个字符串中某个捕获编组的内容
  '示例：Easp.Str.ReplacePart("photo-3.html", "^(\w+)-(\d+)\.html$", "$2", "4")
  '     返回： photo-4.html
  Public Function ReplacePart(ByVal string, ByVal rule, ByVal group, ByVal replaceWith)
    If Not Easp_Test(string, rule) Then
      '如果规则不匹配则直接返回字符串
      ReplacePart = string
      Exit Function
    End If
    Dim o_match, i, j, s_match, i_pos, s_left, s_tmp
    '获取编组号
    i = Int(Mid(group,2))-1
    '取得正则编组
    Set o_match = Match(string,rule)(0)
    '循环编组查找匹配项
    For j = 0 To o_match.SubMatches.Count-1
      s_match = o_match.SubMatches(j)
      '取得当前组的字符开始位置
      i_pos = Instr(string,s_match)
      If i_pos > 0 Then
        '把字符串按当前组的位置分为两部分
        s_tmp = Left(string,i_pos-1)
        string = Mid(string,Len(s_tmp)+1)
        '如果找到匹配的编组号则仅替换本组中的字符串
        If i = j Then
          '把替换后的字符串和前一部分组合起来
          ReplacePart = s_left & s_tmp & o_re.ReFull(string,s_match,replaceWith,i_pos-len(s_tmp),1,0)
          Exit For
        End If
        '如果没有找到匹配则把当前组的字符串换到前一部分中去
        s_left = s_left & s_tmp & s_match
        '在后面部分的字符串中继续下一次扫描匹配
        string = Mid(string, Len(s_match)+1)
      End If
    Next
    Set o_match = Nothing
  End Function
  '正则匹配捕获
  Public Function Match(ByRef string, ByRef rule)
    Dim o_regexp, o_tmp
    Set o_regexp = New Regexp
    o_regexp.Global = True
    o_regexp.IgnoreCase = True
    o_regexp.Pattern = rule
    Set o_tmp = o_regexp.Execute(string)
    Set o_regexp = Nothing
    Set Match = o_tmp
  End Function
  '返回正则验证结果
  Public Function [Test](ByRef string, ByRef rule)
    Dim Pa
    Select Case Lcase(rule)
      Case "date"    [Test] = isDate(string) : Exit Function
      Case "idcard"  [Test] = isIDCard(string) : Exit Function
      Case "number"  [Test] = isNumeric(string) : Exit Function
      Case "english"  Pa = "^[A-Za-z]+$"
      Case "chinese"  Pa = "^[\u4e00-\u9fa5]+$"
      Case "username" Pa = "^[a-zA-Z]\w{2,19}$"
      Case "email"    Pa = "^\w+([-+\.]\w+)*@(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
      Case "int"      Pa = "^[-\+]?\d+$"
      Case "double"   Pa = "^[-\+]?\d+(\.\d+)?$"
      Case "price"    Pa = "^\d+(\.\d+)?$"
      Case "zip"      Pa = "^\d{6}$"
      Case "qq"       Pa = "^[1-9]\d{4,9}$"
      Case "phone"    Pa = "^((\(\+?\d{2,3}\))|(\+?\d{2,3}\-))?(\(0?\d{2,3}\)|0?\d{2,3}-)?[1-9]\d{4,7}(\-\d{1,4})?$"
      Case "mobile"   Pa = "^(\+?\d{2,3})?0?1(3|4|5|7|8)\d{9}$"
      Case "url"      Pa = "^(?:(https|http|ftp|rtsp|mms)://(?:([\w!~\*'\(\).&=\+\$%-]+)(?::([\w!~\*'\(\).&=\+\$%-]+))?@)?)?((?:(?:(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d)\.){3}(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d))|(?:(?:(?:[\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+(?:[a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)|localhost))(?::(\d{1,5}))?([#\?/].*)?$"
      Case "domain"   Pa = "^(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
      Case "ip"       Pa = "^((25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)\.){3}(25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)$"
      Case Else       Pa = rule
    End Select
    [Test] = Easp_Test(string,Pa)
  End Function
  '验证身份证号码
  Private Function isIDCard(ByRef s)
    Dim Ai, BirthDay, arrVerifyCode, Wi, i, AiPlusWi, modValue, strVerifyCode
    isIDCard = False
    If Len(s) <> 15 And Len(s) <> 18 Then Exit Function
    Ai = Easp.IIF(Len(s) = 18,Mid(s, 1, 17),Left(s, 6) & "19" & Mid(s, 7, 9))
    If Not IsNumeric(Ai) Then Exit Function
    If Not Test(Left(Ai,6),"^(1[1-5]|2[1-3]|3[1-7]|4[1-6]|5[0-4]|6[1-5]|8[12]|91)\d{2}[01238]\d{1}$") Then Exit Function
    BirthDay = Mid(Ai, 7, 4) & "-" & Mid(Ai, 11, 2) & "-" & Mid(Ai, 13, 2)
    If IsDate(BirthDay) Then
      If cDate(BirthDay) > Date() Or cDate(BirthDay) < cDate("1870-1-1") Then Exit Function
    Else
      Exit Function
    End If
    arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
    Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
    For i = 0 To 16
      AiPlusWi = AiPlusWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)
    Next
    modValue = AiPlusWi Mod 11
    strVerifyCode = arrVerifyCode(modValue)
    Ai = Ai & strVerifyCode
    If Len(s) = 18 And LCase(s) <> Ai Then Exit Function
    isIDCard = True
  End Function
  '正则替换原型
  Private Function Easp_Replace(ByVal string, ByVal rule, Byval result, ByVal isMultiLine)
    Dim o_regexp
    If Easp.Has(string) Then
      Set o_regexp = New Regexp
      o_regexp.Global = True
      o_regexp.IgnoreCase = True
      If isMultiLine Then o_regexp.Multiline = True
      o_regexp.Pattern = rule
      string = o_regexp.Replace(string,result)
      Set o_regexp = Nothing
    End If
    Easp_Replace = string
  End Function
  '正则匹配原型
  Private Function Easp_Test(ByVal s, ByVal p)
    Dim o_regexp
    If Easp.IsN(s) Then Easp_Test = False : Exit Function
    Set o_regexp = New Regexp
    o_regexp.Global = True
    o_regexp.IgnoreCase = True
    o_regexp.Pattern = p
    Easp_Test = o_regexp.Test(CStr(s))
    Set o_regexp = Nothing
  End Function

  '正则表达式特殊字符转义
  Public Function RegexpEncode(ByVal string)
    Dim re,i
    re = Split("\,$,(,),*,+,.,[,?,^,{,|",",")
    For i = 0 To Ubound(re)
      string = o_re.Re(string, re(i), "\" & re(i))
    Next
    RegexpEncode = string
  End Function

  '将HTML代码转换为文本实体
  Public Function HtmlEncode(ByVal string)
    If Easp.Has(string) Then
      string = o_re.Re(string, Chr(38), "&#38;")
      string = o_re.Re(string, "<", "&lt;")
      string = o_re.Re(string, ">", "&gt;")
      string = o_re.Re(string, Chr(39), "&#39;")
      string = o_re.Re(string, Chr(32), " ")
      string = o_re.Re(string, "  ", "&nbsp;&nbsp;")
      string = o_re.Re(string, Chr(34), "&quot;")
      string = o_re.Re(string, Chr(9), "&nbsp;&nbsp;")
      string = o_re.Re(string, vbCrLf, "<br />")
    End If
    HtmlEncode = string
  End Function
  '将HTML文本转换为HTML代码
  Public Function HtmlDecode(ByVal string)
    If Easp.Has(string) Then
      string = Replace(string, "<br\s*/?\s*>", vbCrLf)
      string = o_re.Re(string, "&nbsp;&nbsp;&nbsp;&nbsp;", Chr(9))
      string = o_re.Re(string, "&quot;", Chr(34))
      string = o_re.Re(string, "&nbsp;", Chr(32))
      string = o_re.Re(string, "&#39;", Chr(39))
      string = o_re.Re(string, "&apos;", Chr(39))
      string = o_re.Re(string, "&gt;", ">")
      string = o_re.Re(string, "&lt;", "<")
      string = o_re.Re(string, "&amp;", Chr(38))
      string = o_re.Re(string, "&#38;", Chr(38))
    End If
    HtmlDecode = string
  End Function

  '过滤HTML标签
  Public Function HtmlFilter(ByVal string)
    If Easp.Has(string) Then
      string = Replace(string,"<[^>]+>","")
      string = o_re.Re(string, ">", "&gt;")
      string = o_re.Re(string, "<", "&lt;")
    End If
    HtmlFilter = string
  End Function
  
  '仅格式化HTML文本中的空格和换行
  Public Function HtmlFormat(ByVal string)
    If Easp.Has(string) Then
      Dim matches, m
      Set matches = Match(string, "<([^>]+)>")
      For Each m In matches
         string = o_re.Re(string, m.SubMatches(0), Replace(m.SubMatches(0), "\s+", Chr(0)))
      Next
      Set matches = Nothing
      string = o_re.Re(string, Chr(32), "&nbsp;")
      string = o_re.Re(string, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
      string = o_re.Re(string, Chr(0), " ")
      string = Replace(string, "(<[^>]+>)\s+", "$1")
      string = o_re.Re(string, vbCrLf, "<br />")
    End If
    HtmlFormat = string
  End Function

  '过滤HTML文本为可输出显示的内容，防止XSS攻击
  'attr: 1-32, 34, 39, 160, 8192-8203, 12288, 65279
  Public Function HtmlSafe(ByVal string)
    If Easp.Has(string) Then
      'string = Asc2Str(string)
      string = Replace(Lcase(string), "<script[\s\S]+?</script\s*>", "")
      string = o_re.Re(string, "<script", "&lt;script")
      string = o_re.Re(string, "</script", "&lt;/script")
      string = o_re.Re(string, "&", "&amp;")
      DropAttrScript string
    End If
    HtmlSafe = string
  End Function
  '替换&#实体字符
  'Private Function Asc2Str(ByVal string)
 '   If Instr(string, "&#") Then
 '     Dim i
 '     For i = 1 To 32
 '       DropAsc string, i
 '     Next
 '     Dim o_matches, m
 '     Set o_matches = Match(string, "&#([a-zA-Z0-9]*);?")
 '     For Each m In o_matches
 '       string = o_re.Re(string, m.Value, ChrW(m.SubMatches(0)))
 '     Next
 '     Set o_matches = Nothing
 '     string = Asc2Str(string)
 '   End If
  '  Asc2Str = string
  'End Function
  '去除无效&#实体字符
  'Private Function DoropAsc(ByRef string, ByVal i)
  '  string = o_re.Re(string, "&#" & i & ";", "")
 '   string = o_re.Re(string, "&#" & i & "", "")
 '   i = Right("0000000" & i, 7)
 '   string = o_re.Re(string, "&#" & i & "", "")
  'End Function
  '去除script标签
  'Private Function DropTagScript(ByVal string)
    'If Test(string, "<script[\s\S]+<")
  'End Function
  '去除属性中的威胁script
  Private Function DropAttrScript(ByRef string)
    Dim o_matches, m, s
    Set o_matches = Match(string, "<[^>]+>")
    For Each m In o_matches
      Easp.PrintHtml m.value
      s = Replace(Lcase(m.value), "[\s""'`]*((j\s*a\s*v\s*a|v\s*b|l\s*i\s*v\s*e)\s*s\s*c\s*r\s*i\s*p\s*t\s*|m\s*o\s*c\s*h\s*a):[^\s""'`>]+", "")
      s = o_re.Re(s, "/*", "")
      s = o_re.Re(s, "*/", "")
      s = Replace(s, ":expression[^;}]+", ":0;")
      Easp.PrintHtml s
      string = o_re.Re(string, m.Value, s)
    Next
    Set o_matches = Nothing
  End Function

  '将对象转换为字符串
  Public Function ToString(ByVal o)
    Dim SB, i, j, k
    Set SB = StringBuilder()
    Select Case VarType(o)
      '如果是数组（可以是多维数组）
      Case 8192,8194,8204,8209
        SB.Append JMultiArray(o)
      Case 8, 9
        '字符串、集合或者对象
        Select Case TypeName(o)
          Case "Connection"
            SB.Append "{""state"":"
            SB.Append o.State
            SB.Append ", ""type"":"""
            SB.Append Easp.Db.GetType(o)
            SB.Append """, ""version"":"""
            SB.Append Easp.Db.GetVersion(o)
            SB.Append """, ""connectionString"":"""
            SB.Append o
            SB.Append """}"
          Case "Recordset"
          '记录集
            If Easp.IsN(o) Then
              SB.Append "{""total"":0, ""rows"":[]}"
            Else
              Set o = o.Clone()
              SB.Append "{""total"":"
              SB.Append o.RecordCount
              SB.Append ", ""rows"":["
              If Not o.BOF Then o.MoveFirst
              i = 0
              Do While Not o.BOF And Not o.EOF
                If i > 0 Then SB.Append ", "
                SB.Append "{"
                For j = 0 To o.Fields.Count-1
                  If j > 0 Then SB.Append ", "
                  SB.Append """"
                  SB.Append o.Fields(j).Name
                  SB.Append """:"
                  If VarType(o.Fields(j).value) = 14 Then
                    SB.Append QuoteString(Trim(o.Fields(j).value))
                  Else
                    SB.Append QuoteString(Easp.IIF(TypeName(o.Fields(j).value)="Byte()", "(blob)", o.Fields(j).value))
                  End If
                Next
                SB.Append "}"
                i = i + 1
                o.MoveNext
              Loop
              o.Close : Set o = Nothing
              SB.Append "]}"
            End If
          Case "Dictionary", "IRequestDictionary", "IReadCookie", "EasyASP_Json_Object", "Errors", "IMatch2"
          '字典对象
            Dim isString
            If TypeName(o) = "IReadCookie" Then
              If o.Count = 0 Then isString = True
            End If
            If isString Then
              SB.Append QuoteString(o)
            Else
              If TypeName(o) = "EasyASP_Json_Object" Then Set o = o.GetDictionary
              SB.Append "{"
              j = 0
              For Each i In o
                If j > 0 Then SB.Append ", "
                SB.Append """"
                SB.Append i
                SB.Append """:"
                SB.Append QuoteString(o(i))
                j = j + 1
              Next
              SB.Append "}"
            End If
          Case "EasyASP_Json_Array"
            o = o.GetArray
            SB.Append JMultiArray(o)
          Case "EasyASP_List"
              SB.Append "{"
              For j = 0 To o.End
                If j > 0 Then SB.Append ", "
                SB.Append """"
                SB.Append o.IndexHash(j)
                SB.Append """:"
                SB.Append QuoteString(o(j))
              Next
              SB.Append "}"
          Case "IMatchCollection2"
          '正则捕获编组
            SB.Append "["
            j = 0
            For Each i In o
              If j > 0 Then SB.Append ", "
              SB.Append "{""match"":"
              SB.Append QuoteString(i.Value)
              If i.SubMatches.Count > 0 Then
                For k = 0 To i.SubMatches.Count - 1
                  SB.Append ",""$"
                  SB.Append k + 1
                  SB.Append """:"
                  SB.Append QuoteString(i.SubMatches(k))
                Next
              End If
              SB.Append "}"
              j = j + 1
            Next
            SB.Append "]"
            '正则搜索集合
            'Case "IMatch2", "Match"
            '  s = FormatReplace(s,t,v.Value)
            '  For i = 0 To v.SubMatches.Count - 1
            '    s = FormatReplace(s,i+t+1,v.SubMatches(i))
            '  Next
          Case "IApplicationObject", "ISessionObject"
          'Application对象和Session对象
            SB.Append "{"
            j = 0
            For Each i In o.Contents
              If j > 0 Then SB.Append ", "
              SB.Append """"
              SB.Append i
              SB.Append """:"
              SB.Append QuoteString(o(i))
              j = j + 1
            Next
            SB.Append "}"
          Case "String", "IStringList"
            '字符串
            If IsNumeric(o) Then o = ToNumber(o, 0)
            SB.Append o
          Case "Nothing"
          Case Else
            SB.Append "{""object"":""unkown"", ""typeName"":"""
            SB.Append TypeName(o)
            SB.Append """, ""varType"":"""
            SB.Append VarType(o)
            SB.Append """}"
        End Select
      Case 0
        SB.Append o
      Case 1
        If TypeName(o) = "Null" Then SB.Append "null"
      Case 2,3,4,5,6
        '数值
        If TypeName(o) = "Object" Then
          'Err对象
          SB.Append "{""number"":"
          SB.Append o.Number
          SB.Append ", ""description"":"
          SB.Append QuoteString(o.Description)
          SB.Append ", ""source"":"
          SB.Append QuoteString(o.Source)
          SB.Append "}"
        Else
          SB.Append ToNumber(o,0)
        End If
      Case 11
        SB.Append Easp.IIF(o, "true", "false")
      Case Else
        SB.Append o
    End Select
    ToString = SB.ToString
    Set SB = Nothing
  End Function
  '返回带引号的字符串
  Private Function QuoteString(ByRef string)
    Dim b_quote
    Select Case VarType(string)
      Case 7
        b_quote = True
      Case 0
        b_quote = True
      Case 8
        'Easp.Console TypeName(string)
        'Easp.Console string
        If IsNumeric(string) Then
          b_quote = False
        ElseIf IsInList("String,IStringList", TypeName(string)) Then
          b_quote = True
        ElseIf TypeName(string) = "IReadCookie" Then
          'Easp.Println string.Count
          If string.Count = 0 Then b_quote = True
          'b_quote = true
        End If
    End Select
    If b_quote Then
      '只有字符串和日期带引号
      QuoteString = """" & JsEncode_(string, b_encodeJson) & """"
    Else
      '其它都不带引号，如是对象再次进行解析
      QuoteString = ToString(string)
    End If
  End Function
  '解析多维数组(based on Jorkin's)
  Private Function JMultiArray(ByRef aArray)
    Dim dimensions, i
    dimensions = getArrayDimension(aArray) '//获取数组维度
    If dimensions > 0 Then
      Dim JMultiArrayExecute, b
      b = "Dim SB " & vbCrLf & "Set SB = StringBuilder()" & vbCrLf
      JMultiArrayExecute = "SB.Append QuoteString(aArray("
      For i = 1 To dimensions
        b = b & "Dim b" & i & vbCrlf '//防止临时变量影响全局变量
        If i > 1 Then JMultiArrayExecute = JMultiArrayExecute & ", "
        JMultiArrayExecute = JMultiArrayExecute & "b" & i
      Next
      JMultiArrayExecute = JMultiArrayExecute & "))" '//生成 aArray(b1, b2, b3, b4....)格式
      For i = 1 To dimensions
        '//一维一维的向外嵌套
        JMultiArrayExecute = "SB.Append ""[""" & vbCrlf & "For b" & i & " = 0 To UBound(aArray, " & i & ")" & vbCrlf & "If b" & i & " > 0 Then SB.Append "", "" End If" & vbCrlf & JMultiArrayExecute & vbCrlf & "Next" & vbCrlf & "SB.Append ""]"""
      Next
      JMultiArrayExecute = JMultiArrayExecute & vbCrLf & "JMultiArray = SB.ToString()" & vbCrLf & "Set SB = Nothing"
      'Easp.Console "<" & "%" & vbCrlf & b & JMultiArrayExecute & vbCrlf & "%" & ">" '//调试生成的语句
      Execute(b & JMultiArrayExecute)
    End If
  End Function
  Private Function getArrayDimension(ByVal aReallydo)
    On Error Resume Next '##Do not delete or comment
    getArrayDimension = 0
    If IsArray(aReallydo) Then
      Dim i, iReallyDo
      For i = 1 To 60
        iReallyDo = UBound(aReallydo, i)
        If Err Then
          Err.Clear
          Exit Function
        Else
          getArrayDimension = i
        End If
      Next
    End If
  End Function
  
  '处理字符串中的Javascript特殊字符，中文使用\uxxxx的形式
  Public Function JsEncode(ByVal string)
    JsEncode = JsEncode_(string, True)
  End Function
  '处理字符串中的Javascript特殊字符
  'cn为False时不处理中文
  Public Function JsEncode_(ByVal string, ByVal cn)
    If Easp.isN(string) Then JsEncode_ = "" : Exit Function
    Dim arr1, arr2, i, j, c, p, SB
    arr1 = Array(&h27,&h22,&h5C,&h2F,&h08,&h0C,&h0A,&h0D,&h09)
    arr2 = Array(&h27,&h22,&h5C,&h2F,&h62,&h66,&h6E,&h72,&h74)
    Set SB = StringBuilder()
    'Easp.Console "::jsencode:" & string
    For i = 1 To Len(string)
      p = True
      c = Mid(string, i, 1)
      For j = 0 To Ubound(arr1)
        If c = Chr(arr1(j)) Then
          SB.Append "\" & Chr(arr2(j))
          p = False
          Exit For
        End If
      Next
      If p Then
        If cn Then
          Dim a
          a = AscW(c)
          If a > 31 And a < 127 Then
            SB.Append c
          ElseIf a > -1 Or a < 65535 Then
            SB.Append "\u" & Right("0000" & Hex(a),4)
          End If
        Else
          SB.Append c
        End If
      End If
    Next
    JsEncode_ = SB.ToString
    Set SB = Nothing
  End Function
  
  '输出javascript代码字符串
  Public Function JavaScript(ByVal string)
    JavaScript = FormatString("<{1} type=""text/java{1}"">{2}{3}{4}{2}</{1}>{2}", Array("sc"&"ript",vbCrLf,vbTab,string),1)
  End Function
  '输出javascript的alert警告框消息
  Public Sub JsAlert(ByVal string)
    Easp.PrintEnd JavaScript(FormatString("alert('{1}');history.go(-1);",JsEncode(string),1))
  End Sub
  '输出javascript的alert警告框消息并跳转到其他页面
  Public Sub JsAlertUrl(ByVal string, ByVal url)
    Easp.PrintEnd JavaScript(FormatString("alert('{1}');location.href='{2}';",Array(JsEncode(string),url),1))
  End Sub
  '输出javascript的选择消息框并根据选择跳转到不同的页面
  Public Sub JsConfirmUrl(ByVal string, ByVal yesUrl, ByVal cancelUrl)
    Easp.PrintEnd JavaScript(FormatString("location.href=confirm('{1}')?'{2}':'{3}';",Array(JsEncode(string),yesUrl,cancelUrl),1))
  End Sub

  '取指定长度的随机字符串
  Public Function RandomStr(ByVal string)
    Dim a, p, l, t, reg, m, mi, ma
    '转义字符
    string = o_re.Re(o_re.Re(o_re.Re(string,"\<",Chr(0)),"\>",Chr(1)),"\:",Chr(2))
    a = ""
    If Easp_Test(string, "(<\d+>|<\d+-\d+>)") Then
    '如果参数中包含 <n> 或 <m-n>
      t = string
      p = GetNameValue(string, ":")
      If Easp.Has(p(1)) Then
        a = p(1) : t = p(0)
      End If
      Set reg = Match(string, "(<\d+>|<\d+-\d+>)")
      For Each m In reg
        p = m.SubMatches(0)
        l = Mid(p,2,Len(p)-2)
        If Easp_Test(l,"^\d+$") Then
        '将包含的所有<n>替换为n位随机字符串
          t = o_re.ReFull(t,p,RandomString(l,a),1,1,0)
        Else
        '将包含的所有<m-n>替换为m到n之间的随机数
          mi = GetName(l,"-")
          ma = GetValue(l,"-")
          t =  o_re.ReFull(t,p,RandomNumber(mi, ma),1,1,0)
        End If
      Next
      Set reg = Nothing
    ElseIf Easp_Test(string,"^\d+-\d+$") Then
    '如果参数为 m-n，则输出m到n之间的随机数
      mi = GetName(string,"-")
      ma = GetValue(string,"-")
      t = RandomNumber(mi, ma)
    ElseIf Easp_Test(string, "^(\d+|\d+:.+)$") Then
    '如果参数为 n 或者 n:string，则输出(string范围中的)n个随机字符串
      l = string : p = GetNameValue(string, ":")
      If Easp.Has(p(1)) Then
        a = p(1) : l = p(0)
      End If
      t = RandomString(l, a)
    Else
      t = string
    End If
    RandomStr = o_re.Re(o_re.Re(o_re.Re(t,Chr(0),"<"),Chr(1),">"),Chr(2),":")
  End Function
  '在指定字符集中取指定长度的随机字符串
  Public Function RandomString(ByVal length, ByVal allowStr)
    Dim i, sb
    If Easp.IsN(allowStr) Then allowStr = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Set sb = StringBuilder()
    For i = 1 To length
      Randomize(Timer)
      sb.Append Mid(allowStr, Int(Len(allowStr) * Rnd + 1), 1)
    Next
    RandomString = sb
    Set sb = Nothing
  End Function

  '得到一个随机数
  Public Function RandomNumber(ByVal min, ByVal max)
    Randomize(Timer) : RandomNumber = Int((max - min + 1) * Rnd + min)
  End Function
  '数字显示指定小数位数，在小于1时显示小数点前面的零
  Public Function ToNumber(ByVal number, ByVal decimalType)
    Dim v, d
    If decimalType < 0 Then
    '如果decimalType为-N，则保留N位小数，但小数位数不足的不补0
      decimalType = 0 - decimalType
      d = True
    ElseIf decimalType = 0 Then
    '如果decimalType为0，则保留所有小数位数
      decimalType = Len(GetValue(CStr(number), "."))
      d = True
    End If
    '如果decimalType为N，则保留N位小数，小数位数不足的补0
    v = FormatNumber(number, decimalType, -1, 0, 0)
    If d And decimalType > 0 Then v = Replace(v, "\.?0+$", "")
    v = Easp.IfHas(v, 0)
    ToNumber = v
  End Function
  '数字显示为货币格式
  Public Function ToPrice(ByVal number)
    ToPrice = FormatCurrency(number, 2, -1, 0, -1)
  End Function
  '数值显示为百分比格式
  Public Function ToPercent(ByVal number)
    ToPercent = FormatPercent(number, 2, -1)
  End Function

  '半角转全角
  Public Function Half2Full(ByVal string)
    'By Demon
    'http://demon.tw
    Dim i
    For i = &H0021 To &H007E
        string = o_re.Re(string, ChrW(i), ChrW(i + &HFEE0))
    Next
    Half2Full = string
  End Function
  '全角转半角
  Public Function Full2Half(ByVal string)
    Dim i
    For i = &HFF01 To &HFF5E
        string = o_re.Re(string, ChrW(i), ChrW(i - &HFEE0))
    Next
    Full2Half = string
  End Function
  '新建一个字符串构建类
  Public Function StringBuilder()
    Set StringBuilder = New EasyASP_Str_StringBuilder
  End Function
  '表单验证
  Public Function Validate(ByVal string)
    Set Validate = New EasyASP_Validation
    Validate.Source = string
  End Function
  '链式操作
  Public Default Function Text(ByVal obj)
    Set Text = New EasyASP_StringObject
    Text.Value = ToString(obj)
  End Function
End Class
'重写VBS内部函数
Class EasyASP_StringOriginal
  Public Function Re(ByVal string, ByVal find, ByVal replacewith)
    Re = Replace(string, find, replaceWith)
  End Function
  Public Function ReCase(ByVal string, ByVal find, ByVal replaceWith)
    ReCase = Replace(string, find, replaceWith, 1, -1, 1)
  End Function
  Public Function ReFull(ByVal string, ByVal find, ByVal replaceWith, ByVal start, ByVal count, ByVal compare)
    ReFull = Replace(string, find, replaceWith, start, count, compare)
  End Function
  Public Function Instr_(ByVal string1, ByVal string2)
    Instr_ = Instr(string1, string2)
  End Function
  Public Function Instr__(ByVal string1, ByVal string2, ByVal start, ByVal compare)
    Instr__ = Instr(start, string1, string2, compare)
  End Function
  Public Function InStrRev_(ByVal string1, ByVal string2)
    InStrRev_ = InStrRev(string1, string2)
  End Function
  Public Function InStrRev__(ByVal string1, ByVal string2, ByVal start, ByVal compare)
    InStrRev__ = InStrRev(string1, string2, start, compare)
  End Function
  Public Function LCase_(ByVal string)
    LCase_ = LCase(string)
  End Function
  Public Function Left_(ByVal string, ByVal length)
    Left_ = Left(string, length)
  End Function
  Public Function Len_(ByVal string)
    Len_ = Len(string)
  End Function
  Public Function LTrim_(ByVal string)
    LTrim_ = LTrim(string)
  End Function
  Public Function RTrim_(ByVal string)
    RTrim_ = RTrim(string)
  End Function
  Public Function Trim_(ByVal string)
    Trim_ = Trim(string)
  End Function
  Public Function Mid_(ByVal string, ByVal start)
    Mid_ = Mid(string, start)
  End Function
  Public Function Mid__(ByVal string, ByVal start, ByVal length)
    Mid__ = Mid(string, start, length)
  End Function
  Public Function Right_(ByVal string, ByVal length)
    Right_ = Right(string, length)
  End Function
  Public Function StrComp_(ByVal string1, ByVal string2)
    StrComp_ = StrComp(string1, string2)
  End Function
  Public Function StrComp__(ByVal string1, ByVal string2, ByVal compare)
    StrComp__ = StrComp(string1, string2, compare)
  End Function
  Public Function StrReverse_(ByVal string)
    StrReverse_ = StrReverse(string)
  End Function
  Public Function Split_(ByVal string, ByVal separator)
    Split_ = Split(string, separator)
  End Function
  Public Function UCase_(ByVal string)
    UCase_ = UCase(string)
  End Function
End Class
%>