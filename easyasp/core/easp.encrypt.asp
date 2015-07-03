<%
'######################################################################
'## easp.encrypt.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Encrypt Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-07-07 3:11:56
'## Description :   Encrypt or decrypt a string in a simple way.
'##
'######################################################################

Class EasyASP_Encrypt
  Private s_seed, s_key
  Private Sub Class_Initialize()
    s_seed = "easyasp_encrypt_seed_by_coldstone"
    s_key = ""
  End Sub
  Public Property Let Key(ByVal string)
    s_key = string
  End Property
  Public Property Get Key()
    Key = s_key
  End Property
  
  Public Default Function Encrypt(ByVal string)
    Encrypt = EncryptBy(string, s_key)
  End Function

  Public Function Decrypt(ByVal string)
    Decrypt = DecryptBy(string, s_key)
  End Function
  
  Public Function EncryptBy(ByVal string, ByVal s_key)
    Dim s_enkey, i_leng, i_keyleng, i_key, i_str, i_re
    Dim SB, i, j, i_times, i_pos, b_flag, s_re
    Dim p1, s1, s2
    If Easp.IsN(string) Then Exit Function
    Set SB = Easp.Str.StringBuilder()
    i_pos = 1
    s_enkey = s_key & s_seed
    i_leng = Len(string)
    i_keyleng = Len(s_enkey)
    i_times = i_leng / i_keyleng
    If i_times > Fix(i_times) Then i_times = Fix(i_times) + 1
    s_re = StrReverse(string)
    If i_leng > 1 Then
      p1 = Fix(i_leng/2)
      s1 = Left(s_re, p1)
      s2 = Mid(s_re, p1+1)
      s_re = s2 & s1
    End If
    For i = 1 To i_times
      For j = 1 To i_keyleng
        i_key = AscW(Mid(s_enkey, j, 1))
        i_str = AscW(Mid(s_re, i_pos, 1))
        i_re = i_str + i_key + 60000
        'Easp.Console i_re & " = " & i_str & " + " & i_key  & " + 42769"
        'Easp.Console n2s(i_re, 43, 39)
        SB.Append n2s(i_re, 50, 39)
        i_pos = i_pos + 1
        If i_pos > i_leng Then
          b_flag = True
          Exit For
        End If
      Next
      If b_flag Then Exit For
    Next
    s_re = SB.ToString
    i_leng = Len(s_re)
    SB.Clear
    For i = 1 To i_leng
      s1 = AscW(Mid(s_re, i, 1))
      If i mod 3 = 1 Then
        s2 = s1
      Else
        s2 = s2 & s1
      End If
      'If Len(s2) = 6 Then Easp.Console s2 & " == " & n2s(s2, 94, 33)
      If Len(s2) = 6 Then SB.Append n2s(s2, 94, 33)
    Next
    EncryptBy = SB.ToString
    Set SB = Nothing
  End Function

  Public Function DecryptBy(ByVal string, ByVal s_key)
    Dim s_enkey, i_leng, i_keyleng, i_key, i_str, i_re
    Dim SB, i, j, i_times, i_pos, b_flag, s_re
    Dim p1, s1, s2, Ar, a_re
    If Easp.IsN(string) Then Exit Function
    Set Ar = Easp.Json.NewArray()
    i_leng = Len(string)
    For i = 1 To i_leng Step 3
      s1 = Mid(string, i, 3)
      s2 = s2n(s1, 94, 33)
      s_re = ""
      For j = 1 To 6 Step 2
        s_re = s_re & ChrW(Mid(s2, j, 2))
      Next
      'Easp.Console s2 & " == " & s_re & " -- " & s2n(s_re, 50, 39)
      Ar.Add s2n(s_re, 50, 39)
    Next
    's_re = SB.ToString
    a_re = Ar.GetArray()
    Set Ar = Nothing
    i_pos = 0
    s_enkey = s_key & s_seed
    i_leng = Ubound(a_re) + 1
    i_keyleng = Len(s_enkey)
    i_times = i_leng / i_keyleng
    If i_times > Fix(i_times) Then i_times = Fix(i_times) + 1
    Set SB = Easp.Str.StringBuilder()
    For i = 1 To i_times
      For j = 1 To i_keyleng
        i_key = AscW(Mid(s_enkey, j, 1))
        i_str = a_re(i_pos)
        i_re = i_str - i_key - 60000
        'Easp.Console Mid(s_enkey, j, 1) & " - "  & i_pos
        'Easp.Console i_re & " = " & i_str & " - " & i_key & " - 42769"
        'Easp.Println i_re
        SB.Append ChrW(i_re)
        i_pos = i_pos + 1
        If i_pos >= i_leng Then
          b_flag = True
          Exit For
        End If
      Next
      If b_flag Then Exit For
    Next
    s_re = SB.ToString
    Set SB = Nothing
    i_leng = Len(s_re)
    If i_leng > 1 Then
      p1 = Fix(i_leng/2)
      s1 = Right(s_re, p1)
      s2 = Left(s_re, i_leng-p1)
      s_re = s1 & s2
    End If
    DecryptBy = StrReverse(s_re)
  End Function

  Private Function n2s(ByVal n, ByVal sys, ByVal start)
    Dim t(2), v, c, m, l
    c = 3
    v = n / sys
    Do While v > 0 And c > 0
      c = c - 1
      m = Int(n - Int(n / sys) * sys)
      t(c) = ChrW(m+start)
      n = v
      v = n / sys
    Loop
    n2s = Join(t,"")
  End Function
  Private Function s2n(ByVal s, ByVal sys, ByVal start)
    Dim i, t, m, n
    m = 1
    n = 0
    For i = 1 to 3
      t = Mid(s, i, 1)
      m = AscW(t)-start
      n = n + sys^(3-i) * m
    Next
    s2n = n
  End Function
End Class
%>