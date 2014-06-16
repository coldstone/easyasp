<%
'######################################################################
'## easp.encrypt.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Encrypt Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-16 15:07:39
'## Description :   Encrypt or decrypt a string in a simple way.
'##
'######################################################################

Class EasyASP_Encrypt
  Private s_seed, s_key
  Private Sub Class_Initialize()
    s_seed = "easyaspencryptbycoldstone"
    s_key = ""
  End Sub
  Public Property Let Key(ByVal string)
    s_key = string
  End Property
  Public Property Get Key()
    Key = s_key
  End Property
  Public Default Function Encrypt(ByVal string)
    Encrypt = Crypting(string, True)
  End Function
  Public Function Decrypt(ByVal string)
    Decrypt = Crypting(string, False)
  End Function
  Private Function Crypting(ByVal string, ByVal isEnCrypt)
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
    If Not isEnCrypt Then
      If i_leng > 1 Then
        p1 = Fix(i_leng/2)
        s1 = Right(string, p1)
        s2 = Left(string, i_leng-p1)
        string = s1 & s2
      End If
      string = StrReverse(string)
    End If
    For i = 1 To i_times
      For j = 1 To i_keyleng
        i_key = AscW(Mid(s_enkey, j, 1))
        i_str = AscW(Mid(string, i_pos, 1))
        i_re = i_str + Easp.IIF(isEnCrypt, i_key, 0-i_key)
        SB.Append ChrW(i_re)
        i_pos = i_pos + 1
        If i_pos > i_leng Then
          b_flag = True
          Exit For
        End If
      Next
      If b_flag Then Exit For
    Next
    If isEnCrypt Then
      s_re = StrReverse(SB.ToString)
      If i_leng > 1 Then
        p1 = Fix(i_leng/2)
        s1 = Left(s_re, p1)
        s2 = Mid(s_re, p1+1)
        s_re = s2 & s1
      End If
    Else
      s_re = SB.ToString
    End If
    Crypting = s_re
    Set SB = Nothing
  End Function
End Class
%>