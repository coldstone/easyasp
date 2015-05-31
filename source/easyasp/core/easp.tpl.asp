<%
'######################################################################
'## easp.tpl.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Templates Class
'## Version     :   v3
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-07-25 10:08:34
'## Description :   Use Templates with EasyASP
'##
'######################################################################
Class EasyASP_Tpl
  Private s_html, s_ohtml, s_unknown, s_dict, s_path, s_m, s_ms, s_me
  Private o_tag, o_blockdata, o_block, o_blocktag, o_blocks, o_attr
  Private b_asp

  Private Sub class_Initialize
    s_path = ""
    s_unknown = "keep"
    s_dict = "Scripting.Dictionary"
    Set o_tag = Server.CreateObject(s_dict) : o_tag.CompareMode = 1
    Set o_blockdata = Server.CreateObject(s_dict) : o_blockdata.CompareMode = 1
    Set o_block = Server.CreateObject(s_dict) : o_block.CompareMode = 1
    Set o_blocktag = Server.CreateObject(s_dict) : o_blocktag.CompareMode = 1
    Set o_blocks = Server.CreateObject(s_dict) : o_blocks.CompareMode = 1
    Set o_attr = Server.CreateObject(s_dict) : o_attr.CompareMode = 1
    s_m = "{*}"
    getMaskSE s_m
    b_asp = False
    s_html = ""
    s_ohtml = ""
  End Sub
  Private Sub Class_Terminate
    Set o_tag = Nothing
    Set o_blockdata = Nothing
    Set o_block = Nothing
    Set o_blockTag = Nothing
    Set o_blocks = Nothing
    Set o_attr = Nothing
  End Sub

  '设置和读取静态模板文件路径
  Public Property Get FilePath
    FilePath = s_path
  End Property
  Public Property Let FilePath(ByVal f)
    If Right(f,1)<>"/" Then f = f & "/"
    s_path = f
  End Property
  '加载模板文件
  Public Property Let [File](ByVal f)
    Load(f)
  End Property
  '通过文本加载模板
  Public Property Let [Source](ByVal s)
    LoadStr(s)
  End Property
  '设置和读取标签的标识符
  Public Property Get TagMask
    TagMask = s_m
  End Property
  Public Property Let TagMask(ByVal m)
    s_m = m
    getMaskSE s_m
  End Property
  '标签替换（属性模式）
  Public Property Let Tag(ByVal s, ByVal v)
    Assign s, v
  End Property
  Public Property Get Tag(ByVal s)
    If o_tag.Exists(s) Then
      Tag = o_tag.Item(s)
    Else
      Tag = ""
    End If
  End Property
  '设置模板中是否可以执行ASP代码
  Public Property Get AspEnable
    AspEnable = b_asp
  End Property
  Public Property Let AspEnable(ByVal b)
    b_asp = b
  End Property
  '设置如何处理未定义的标签
  Public Property Get TagUnknown
    TagUnknown = s_unknown
  End Property
  Public Property Let TagUnknown(ByVal s)
    Select Case LCase(s)
      Case "1", "remove"
        s_unknown = "remove"
      Case "2", "comment"
        '转成注释慎用，如果是html标签的属性值内有标签转为注释可能引发html标签错误
        s_unknown = "comment"
      Case Else
        s_unknown = "keep"
    End Select
  End Property
  '建新Easp模板类实例
  Public Function [New]()
    Set [New] = New EasyASP_Tpl
  End Function
  '读取循环块的属性
  Public Function Attr(ByVal s)
    If Not o_attr.Exists(s) Then Exit Function
    Attr = o_attr.Item(s)
  End Function

  '加载静态模板文件
  Public Sub Load(ByVal f)
    s_html = LoadInc(s_path & f,"")
    s_ohtml = s_html
    SetBlocks()
  End Sub
  '从文本加载模板
  Public Sub LoadStr(ByVal s)
    s_html = s
    s_ohtml = s
    SetBlocks()
  End Sub
  '重新载入当前模板
  Public Sub Reload()
    o_tag.RemoveAll
    o_blockdata.RemoveAll
    o_block.RemoveAll
    o_blockTag.RemoveAll
    o_blocks.RemoveAll
    o_attr.RemoveAll
    s_html = s_ohtml
    SetBlocks()
  End Sub
  '载入附加模板到标签
  Public Sub TagFile(ByVal tag, ByVal f)
    LoadToTag tag,0,f
  End Sub
  '从文本加载附加模板到标签
  Public Sub TagStr(ByVal tag, ByVal s)
    LoadToTag tag,1,s
  End Sub
  '加载附加模板原型
  Private Sub LoadToTag(ByVal tag, ByVal t, ByVal f)
    Dim s
    If t = 0 Then
      s = LoadInc(s_path & f,"")
    Else
      s = f
    End If
    If Easp.Has(tag) Then
      s_html = Easp.Str.Replace(s_html, s_ms & tag & s_me, s)
    Else
      s_html = s_html & s
    End If
    SetBlocks()
  End Sub
  '替换标签（默认方法）
  Public Default Sub Assign(ByVal s, ByVal v)
    Dim i,f
    Select Case TypeName(v)
      '替换记录集标签
      Case "Recordset"
        If Easp.Has(v) Then
          For i = 0 To v.Fields.Count - 1
            Assign s & "." & v.Fields(i).Name, Trim(v.Fields(i).Value)
            Assign s & "." & i, Trim(v.Fields(i).Value)
          Next
        End If
      '替换Easp超级数组标签
      Case "EasyASP_List"
        If v.Size > 0 Then
          For i = 0 To v.End
            Assign s & "." & i, v(i)
            Assign s & "." & v.IndexHash(i), v(i)
          Next
        End If
      Case Else
        If Easp.IsN(v) Then v = ""
        If o_tag.Exists(s) Then o_tag.Remove s
        o_tag.Add s, Cstr(v)
    End Select
  End Sub
  '在已替换标签后添加新内容
  Public Sub Append(ByVal s, ByVal v)
    If Easp.IsN(v) Then v = ""
    Dim tmp
    If o_tag.Exists(s) Then
      tmp = o_tag.Item(s) & Cstr(v)
      o_tag.Remove s
      o_tag.Add s, Cstr(tmp)
    Else
      o_tag.Add s, Cstr(v)
    End If
  End Sub
  '更新循环块数据
  Public Sub [Update](ByVal b)
    Dim Matches, Match, tmp, s, rule, data
    s = BlockData(b)
    rule = Chr(0) & "(\w+?)" & Chr(0)
    Set Matches = Easp.Str.Match(s, rule)
    Set Match = Matches
    For Each Match In Matches
      data = Match.SubMatches(0)
      If o_blocktag.Exists(data) Then
        s = Replace(s, Match.Value, o_blocktag.Item(data))
        o_blocktag.Remove(data)
      End If
    Next
    If o_blocktag.Exists(b) Then
      tmp = o_blocktag.Item(b) & s
      o_blocktag.Remove b
      o_blocktag.Add b, Cstr(tmp)
    Else
      o_blocktag.Add b, Cstr(s)
    End If
    Set Matches = Easp.Str.Match(s_html, Chr(0) & b & Chr(0))
    Set Match = Matches
    For Each Match In Matches
      s = BlockTag(b)
      s_html = Replace(s_html, Match.Value, s & Match.Value)
    Next
    If o_block.Exists(b) Then o_block.Remove b
  End Sub
  '处理逻辑控制块(if..else)
  Private Function LogicReplace(ByVal s)
    Dim Matches, Match, result, condi, conds, cond, n, yes, no, x, e, f, cname, ckey, copera, cvalue
    Set Matches = Easp.Str.Match(s, s_ms & "#if\s+(.+?)"&s_me&"([\s\S]+?)(?:"&s_ms&"#else"&s_me&"([\s\S]+?))?"&s_ms&"/#if"&s_me)
    For Each Match In Matches
      condi = Match.SubMatches(0)
      yes = Match.SubMatches(1)
      no = Match.SubMatches(2)
      '选择条件表达式组
      Set conds = Easp.Str.Match(condi,"(?:([^&)(\s}|}=<>!]+)([=<>!]{1,2})(['""])(.+?)\3)|(?:([^&()\s|}=<>!]+)([=<>!]{1,2})([^&)\s|}]+))")
      For Each cond In conds
        '选择到的表达式
        cname = cond.Value
        '找其中的变量标签
        Set e = Easp.Str.Match(cname,"@([\w\.]+)")
        For Each f In e
          n = f.SubMatches(0)
          '把标签替换为值
          cname = Replace(cname, f.Value, Easp.IIF(o_tag.Exists(n),o_tag.Item(n),""))
        Next
        Set e = Nothing
        '解析表达式
        Set x = Easp.Str.Match(cname,"^([^=<>!]*)([=<>!]{1,2})(['""]?)(.*)\3$")
        ckey = x(0).SubMatches(0)
        copera = x(0).SubMatches(1)
        cvalue = x(0).SubMatches(3)
        'Easp.WNH "exp:" & ckey & copera & cvalue
        '比较表达式的结果
        condi = Replace(condi, cond.Value, Comp(ckey,copera,cvalue))
        condi = Replace(condi, "&&", " And ")
        condi = Replace(condi, "||", " Or ")
        Set x = Nothing
      Next
      'Easp.WNH condi
      'Easp.WN yes
      'Easp.WN no
      s = Replace(s, Match.Value, Easp.IIF(Eval(condi), yes, no))
      Set conds = Nothing
    Next
    Set Matches = Nothing
    LogicReplace = s
  End Function
  '比较表达式(感谢Taihom)
  Private Function Comp(ByVal k, ByVal o, ByVal v)
    On Error Resume Next '##Do not delete or comment
    Dim tmp,m,ma,mb : tmp = False
    Select Case o
      Case "=","=="
        m = Replace(k,"\%","")
        If Instr(m,"%")>0 Then
          ma = Easp.Str.GetName(m,"%")
          mb = Easp.Str.GetValue(m,"%")
          tmp = (CLng(ma) Mod CLng(mb) = v)
        Else
          tmp = (CStr(k) = CStr(v))
        End If
      Case "<>","!=" tmp = (CStr(k) <> CStr(v))
      Case ">=" tmp = (CDbl(k) >= CDbl(v))
      Case "<=" tmp = (CDbl(k) <= CDbl(v))
      Case ">" tmp = (CDbl(k) > CDbl(v))
      Case "<" tmp = (CDbl(k) < CDbl(v))
    End Select
    Comp = Easp.IIF(Err.Number=0,tmp,False)
  End Function
  '获取最终html
  Public Function GetHtml()
    s_html = LogicReplace(s_html)
    Dim Matches, Match, n, f, f0, b
    '替换标签
    Set Matches = Easp.Str.Match(s_html, s_ms & "([^:]+?)?(:.+?)?" & s_me)
    'Easp.WN "rule:" & s_ms & "(.+?)" & s_me
    For Each Match In Matches
      n = Match.SubMatches(0)
      f = Match.SubMatches(1)
      f0 = Easp.IIF(Easp.Has(f),"{0"&f&"}","{0}")
      'Easp.WT f0
      'Easp.WN "match://" & Match.Value & "//"
      If o_tag.Exists(n) Then
        'Easp.WT f0
        s_html = Replace(s_html, Match.Value, Easp.Str.Format(f0,o_tag.Item(n)))
        's_html = Replace(s_html, Match.Value, o_tag.Item(n))
        'Easp.WN "match_tag:" & Match.Value
        'Easp.WN "match_dic:" & o_tag.Item(n)
      End If
    Next
    '替换未处理循环块
    Set Matches = Easp.Str.Match(s_html, Chr(0) & "(\w+?)" & Chr(0))
    For Each Match In Matches
      b = Match.SubMatches(0)
      Select Case s_unknown
        Case "keep"
          If o_block.Exists(b) Then [Update](b)
        Case "remove"
          'Do Nothing
        Case "comment"
          s_html = Replace(s_html, Match.Value, "<!-- Unknown Block '"&b&"' -->")
      End Select
      s_html = Replace(s_html, Match.Value, "")
    Next
    '替换未处理标签
    Set Matches = Easp.Str.Match(s_html, s_ms & "(.+?)" & s_me)
    select case s_unknown
      case "keep"
        'Do Nothing
      case "remove"
        For Each Match In Matches
          s_html = Replace(s_html, Match.Value, "")
        Next
      case "comment"
        For Each Match In Matches
          s_html = Replace(s_html, Match.Value, "<!-- Unknown Tag '" & Match.Submatches(0) & "' -->")
        Next
    End select
    Set Matches = Nothing
    GetHtml = s_html
  End Function
  '输出最终文件内容
  Public Sub Show()
    Easp.Print GetHtml
  End Sub
  '保存为静态文件
  Public Sub SaveAs(ByVal p)
    Call Easp.Fso.CreateFile(p,GetHtml)
  End Sub
  '生成html标签
  Public Function MakeTag(ByVal t, ByVal f)
    Dim s,e,a,i,m
    If Instr(t,":")>0 Then
      m = Easp.Str.GetValue(t,":")
      t = Easp.Str.GetName(t,":")
      m = Easp.Date.Format(Now,m)
    End If
    Select Case Lcase(t)
      Case "css"
        s = "<link href="""
        e = """ rel=""stylesheet"" type=""text/css"" />"
      Case "js"
        s = "<scr"&"ipt type=""text/javas"&"cript"" src="""
        e = """></scr"&"ipt>"
      Case "author", "keywords", "description", "copyright", "generator", "revised", "others"
        MakeTag = MakeTagMeta("name",t,f)
        Exit Function
      Case "content-type", "expires", "refresh", "set-cookie"
        MakeTag = MakeTagMeta("http-equiv",t,f)
        Exit Function
    End Select
    a = Split(f,"|")
    For i = 0 To Ubound(a)
      a(i) = s & Trim(a(i)) & Easp.IfThen(Easp.Has(m),"?" & m) & e
    Next
    MakeTag = Join(a,vbCrLf)
  End Function

  '生成Meta标签
  Private Function MakeTagMeta(ByVal m, ByVal t, ByVal s)
    MakeTagMeta = "<meta " & m & "=""" & t & """ content=""" & s & """ />"
  End Function
  '获取Tag标识
  Private Sub getMaskSE(ByVal m)
    s_ms = Easp.Str.RegexpEncode(Easp.Str.GetName(m,"*"))
    s_me = Easp.Str.RegexpEncode(Easp.Str.GetValue(m,"*"))
  End Sub
  '载入模板文件并将无限级include模板载入
  Private Function LoadInc(ByVal f, ByVal p)
    Dim h,pa,rule,inc,Match,incFile,incStr
    pa = Easp.IIF(Left(f,1)="/","",p)
    If b_asp Then
      h = Easp.GetInclude( pa & f )
    Else
      h = Easp.Fso.Read( pa & f )
    End If
    rule = "(<!--[\s]*)?" & s_ms & "#include:(.+?)" & s_me & "([\s]*-->)?"
    If Easp.Str.Test(h,rule) Then
      If Easp.isN(p) Then
        If Instr(f,"/")>0 Then p = Left(f,InstrRev(f,"/"))
      Else
        If Instr(f,"/")>0 Then p = pa & Left(f,InstrRev(f,"/"))
      End If
      Set inc = Easp.Str.Match(h,rule)
      For Each Match In inc
        incFile = Match.SubMatches(1)
        incStr = LoadInc(incFile, p)
        h = Replace(h,Match,incStr)
      Next
      Set inc = Nothing
    End If
    LoadInc = h
  End Function
  '读取循环块标签
  Private Sub SetBlocks()
    Dim Matches, Match, rule, n, i, j
    i = 0
    rule = "(<!--[\s]*)?" & s_ms & "#:(.+?)" & s_me
    If Not Easp.Str.Test(s_html, rule) Then Exit Sub
    Set Matches = Easp.Str.Match(s_html,rule)
    '找到循环块
    For Each Match In Matches
      n = Match.SubMatches(1)
      'Easp.WN "block:" & n
      '把循环块标签加入字典
      If o_blocks.Exists(i) Then o_blocks.Remove i
      o_blocks.Add i, n
      i = i + 1
    Next
    '从最后一层开始初始化循环块，实现无限层嵌套
    For j = i-1 To 0 Step -1
      Begin o_blocks.item(j)
    Next
  End Sub
  '初始化循环块
  Private Sub Begin(ByVal b)
    Dim Matches, Match, rule, data, attrs, attr, att, aname, avalue, atag
    rule = "(<!--[\s]*)?(" & s_ms & ")#:(" & b & ")(" & s_me & ")([\s]*-->)?([\s\S]+?)(<!--[\s]*)?\2/#:\3\4([\s]*-->)?"
    '如果循环块有属性则取出属性
    If Instr(b," ")>0 Then
      attrs = Easp.Str.GetValue(b, " ")
      b = Easp.Str.GetName(b, " ")
      rule = "(<!--[\s]*)?(" & s_ms & ")#:(" & b & " " & Easp.Str.RegexpEncode(attrs) & ")(" & s_me & ")([\s]*-->)?([\s\S]+?)(<!--[\s]*)?\2/#:" & b & "\4([\s]*-->)?"
    End If
    Set Matches = Easp.Str.Match(s_html, rule)
    Set Match = Matches
    For Each Match In Matches
      'Easp.WN "block_tag:" & b
      'Easp.WN "block_attr:" & attrs
      '取循环块内容
      data = Match.SubMatches(5)
      '把循环块内容存入标签名对应的字典
      If o_blockdata.Exists(b) Then
        o_blockdata.Remove(b)
        o_block.Remove(b)
      End If
      o_blockdata.Add b, Cstr(data)
      o_block.Add b, Cstr(b)
      If Easp.Has(attrs) Then
      '如果有属性则取出每个属性
        Set attr = Easp.Str.Match(attrs,"((\w+)=(['""])(.+?)\3)|((\w+)=([^\s]+))")
        For Each att In attr
          aname = Easp.Str.GetName(att.Value, "=")
          avalue = Easp.Str.Replace(att.Value, "\w+=(['""]?)(.+?)\1", "$2")
          atag = b & "." & aname
          'Easp.WN "attr '" & aname & "' = ["&avalue&"]"
          If o_attr.Exists(atag) Then o_attr.Remove(atag)
          o_attr.Add atag, avalue
        Next
      End If
      '把原始内容中的循环块作临时替换
      s_html = Easp.Str.Replace(s_html, rule, Chr(0) & b & Chr(0))
    Next
    Set Matches = Nothing
  End Sub
  '取循环块原始模板数据
  Private Function BlockData(ByVal b)
    Dim tmp, s
    If o_blockdata.Exists(b) Then
      tmp = o_blockdata.Item(b)
      '替换已定义标签
      s = UpdateBlockTag(tmp)
      BlockData = s
    Else
      BlockData = "<!--" & Chr(0) & b & Chr(0) & "-->"
    End If
  End Function
  '取循环块临时数据
  Private Function BlockTag(ByVal b)
    Dim tmp, s
    If o_blockdata.Exists(b) Then
      tmp = o_blocktag.Item(b)
      '替换已定义标签
      s = UpdateBlockTag(tmp)
      BlockTag = s
      '删除循环块临时数据
      o_blocktag.Remove(b)
    Else
      BlockTag = "<!--" & Chr(0) & b & Chr(0) & "-->"
    End If
  End Function
  '更新循环块标签
  Private Function UpdateBlockTag(ByVal s)
    s = LogicReplace(s)
    Dim Matches, Match, data, rule, f, f0
    Set Matches = Easp.Str.Match(s, s_ms & "([^:]+?)?(:.+?)?" & s_me)
    For Each Match In Matches
      '取标签名
      data = Match.SubMatches(0)
      f = Match.SubMatches(1)
      f0 = Easp.IIF(Easp.Has(f),"{0"&f&"}","{0}")
      '如果此标签有替换值
      If o_tag.Exists(data) Then
        rule = Match.Value      
        '替换标签为相应的值
        If Easp.isN(o_tag.Item(data)) Then
          s = Replace(s, rule, "")
        Else
          s = Replace(s, rule, Easp.Str.Format(f0,o_tag.Item(data)))
        End If
      End If
    Next
    UpdateBlockTag = s
  End Function
End Class
%>