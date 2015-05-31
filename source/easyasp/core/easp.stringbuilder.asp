<%
'######################################################################
'## easp.stringbuilder.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP StringBuilder Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-05-12 0:38:21
'## Description :   Create a string in a high-performance way
'##
'######################################################################

'字符串构造类
Class EasyASP_Str_StringBuilder
  Private a_sb(), i_index, a_sbi(), i_indexi
  Private i_length, b_line, b_insert
  Private Sub Class_Initialize()
    i_index  = 0
    i_indexi = 99
    i_length = 99
    ReDim a_sb(i_length)
    ReDim a_sbi(i_length)
    b_line = False
    b_insert = False
  End Sub
  Private Sub Class_Terminate()
  End Sub
  '是否附加为新行
  Public Property Let NewLine(ByVal bool)
    b_line = bool
  End Property
  '设置容量
  Public Property Let Capacity(ByVal number)
    i_length = number - 1
    ReDim a_sb(i_length)
  End Property
  '返回当前容量
  Public Property Get Capacity
    Capacity = i_length + 1
  End Property
  
  '附加字符串
  Public Sub Append(ByVal string)
    AppendString string, False, ""
  End Sub
  '以新行方式附加字符串
  Public Sub AppendLine(ByVal string)
    AppendString string, True, ""
  End Sub
  '带格式化附加字符串
  Public Sub AppendFormat(ByVal string, ByVal format)
    AppendString string, False, format
  End Sub
  '附加字符串原型
  Private Sub AppendString(ByVal string, ByVal newLine, ByVal format)
    Dim s_tmp, b_format
    If i_index >= i_length Then
      s_tmp = Join(a_sb, "")
      ReDim a_sb(i_length)
      a_sb(0) = s_tmp
      i_index = 1
    End If
    If IsArray(format) Or IsObject(format) Then
      b_format = True
    ElseIf format > "" Then
      b_format = True
    End If
    If b_format Then
      a_sb(i_index) = Easp.Str.Format(string, format)
    Else
      a_sb(i_index) = string
    End If
    i_index = i_index + 1
    If newLine Or b_line Then
      a_sb(i_index) = vbCrLf
      i_index = i_index + 1
    End If
  End Sub

  '从开始处插入字符串
  Public Sub Insert(ByVal string)
    InsertString string, False, ""
  End Sub
  '以新行方式从开始处插入字符串
  Public Sub InsertLine(ByVal string)
    InsertString string, True, ""
  End Sub
  '从开始处插入带格式化字符串
  Public Sub InsertFormat(ByVal string, ByVal format)
    InsertString string, True, format
  End Sub
  '从开始处插入字符串原型
  Private Sub InsertString(ByVal string, ByVal newLine, ByVal format)
    Dim s_tmp, b_format
    If i_indexi <= 0 Then
      s_tmp = Join(a_sbi, "")
      ReDim a_sbi(i_length)
      a_sbi(i_length) = s_tmp
      i_indexi = i_length - 1
    End If
    If newLine Or b_line Then
      a_sbi(i_indexi) = vbCrLf
      i_indexi = i_indexi - 1
    End If
    If IsArray(format) Or IsObject(format) Then
      b_format = True
    ElseIf format > "" Then
      b_format = True
    End If
    If b_format Then
      a_sbi(i_indexi) = Easp.Str.Format(string, format)
    Else
      a_sbi(i_indexi) = string
    End If
    i_indexi = i_indexi - 1
    If Not b_insert Then b_insert = True
  End Sub
  '清除所有字符
  Public Sub Clear()
    ReDim a_sb(i_length)
    If b_insert Then ReDim a_sbi(i_length)
  End Sub
  '输出字符串
  Public Default Function ToString()
    If b_insert Then
      ToString = Join(Array(Join(a_sbi, ""), Join(a_sb, "")), "")
    Else
      ToString = Join(a_sb, "")
    End If
  End Function
End Class
%>