<%
'######################################################################
'## easp.hanzi.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Chinese character processing tools
'## Version     :   1.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-04-16 11:31:25
'## Description :   This plugin provides Chinese characters to Pinyin,
'##                 Chinese characters translate English , Chinese word
'##                 segmentation , etc.
'##
'######################################################################
Class EasyASP_Hanzi
  Private s_cn, s_author, s_version
  Private b_title
  Private a_result(2)

  Private Sub Class_Initialize()
		s_author	= "coldstone"
		s_version	= "1.0"
    b_title = True
  End Sub
  Public Property Get Author()
    Author = s_author
  End Property
  Public Property Get Version()
    Version = s_version
  End Property

  '取得的拼音结果每个字的首字母是否大写
  Public Property Let TitleCase(ByVal bool)
    b_title = bool
  End Property
  Public Property Get TitleCase()
    TitleCase = b_title
  End Property
  
  '取汉字拼音
  Public Function GetPinYin(ByRef chinese)
    GetPinYin = ClearPinYin(GetResults(chinese)(0), True, False, False, False, b_title)
  End Function

  '取汉字拼音首字母
  Public Function GetPY(ByRef chinese)
    GetPY = ClearPinYin(GetResults(chinese)(0), True, False, False, True, b_title)
  End Function

  '取带声调的汉字拼音
  Public Function GetPinYinRead(ByRef chinese)
    GetPinYinRead = GetResults(chinese)(0)
  End Function

  '取带声调的汉字拼音，声调以1234标识
  Public Function GetPinYin1234(ByRef chinese)
    GetPinYin1234 = ClearPinYin(GetResults(chinese)(0), True, True, True, False, b_title)
  End Function

  '自定义取拼音的结果样式
  '参数：("中文字符串", 拼音韵母转为字母, 拼音后标识声调, 拼音间加空格, 仅取首字母, 首字母大写)
  Public Function GetPinYinWith(ByVal chinese, ByVal toneToLetter,_
                                ByVal hasToneNumber, ByVal hasSpace,_
                                ByVal onlyFirstLetter, ByVal Title)
    GetPinYinWith = ClearPinYin(GetResults(chinese)(0), toneToLetter, hasToneNumber, hasSpace, onlyFirstLetter, Title)
  End Function

  '汉字翻译为英文
  Public Function GetEnglish(ByRef chinese)
    GetEnglish = GetResults(chinese)(1)
  End Function

  '取得翻译后的英文并以短横线（-）分隔
  Public Function GetEnglishDash(ByRef chinese)
    Dim s_result
    s_result = GetResults(chinese)(1)
    s_result = Easp.Str.Replace(s_result, "[,!,.。，、；：？！…—\·ˉˇ¨\/\\\?;'""\:\[\]\{\}\|\-_\+=~@#\$\%\^&\*\(\)]", "")
    s_result = LCase(Trim(Easp.Str.Replace(s_result, "\s+", " ")))
    GetEnglishDash = Join(Split(s_result), "-")
  End Function

  '汉字分词关键字，分词结果以空格分隔
  Public Function GetKeyWord(ByRef chinese)
    GetKeyWord = Join(GetResults(chinese)(2), " ")
  End Function

  '汉字分词关键字，分词结果为数组
  Public Function GetKeyWordArray(ByRef chinese)
    GetKeyWordArray = GetResults(chinese)(2)
  End Function

  '清除拼音结果中的多余字符，并按条件生成结果
  '参数：("中文字符串", 拼音韵母转为字母, 拼音后标识声调, 拼音间加空格, 仅取首字母, 首字母大写)
  Private Function ClearPinYin(ByVal string, ByVal toneToLetter,_
                               ByVal hasToneNumber, ByVal hasSpace,_
                               ByVal onlyFirstLetter, ByVal Title)
    string = Easp.Str.Replace(string, "([,!,.。，、；：？！…—\·ˉˇ¨\/\\\?;'""“”\:\[\]\{\}\|\-_\+=~@#\$\%\^&\*\(\)])", " ")
    string = Easp.Str.Replace(string, "ng(b|p|m|f|d|t|l|n|g|k|h|j|q|x|r|y|w|z|c|s|\s+|[,\/\\\}\{""'\:;+-_=.!@#$%\[\]&\*\(\)。，、；：？！…—·ˉˇ¨‘’“”‖《》〉〈＂＇｀｜〃〔〕「」『』．〖〗【】（）［］｛｝]|$)", "{**}$1")
    string = Easp.Str.Replace(string, "n(b|p|m|f|d|t|l|n|g|k|h|j|q|x|r|y|w|z|c|s|\s+|[,\/\\\}\{""'\:;+-_=.!@#$%\[\]&\*\(\)。，、；：？！…—·ˉˇ¨‘’“”‖《》〉〈＂＇｀｜〃〔〕「」『』．〖〗【】（）［］｛｝]|$)", "{*}$1")
    string = Easp.Str.Replace(string, "zh", "{*1}")
    string = Easp.Str.Replace(string, "ch", "{*2}")
    string = Easp.Str.Replace(string, "sh", "{*3}")
    string = Easp.Str.Replace(string, "(b|p|m|f|d|t|l|n|g|k|h|j|q|x|\{\*1\}|\{\*2\}|\{\*3\}|r(?=[aiouev])|y|w|z|c|s)|\s+", " $1")
    string = Easp.Str.Replace(string, "\{\*\}", "n")
    string = Easp.Str.Replace(string, "\{\*\*\}", "ng")
    string = Easp.Str.Replace(string, "\{\*1\}", "zh")
    string = Easp.Str.Replace(string, "\{\*2\}", "ch")
    string = Easp.Str.Replace(string, "\{\*3\}", "sh")
    string = Easp.Str.Replace(string, "\s+", " ")
    string = Trim(LCase(string))
    Dim a_letter, i
    a_letter = Split(string)
    For i = 0 To UBound(a_letter)
      a_letter(i) = SwitchTone(a_letter(i), toneToLetter, hasToneNumber)
      If Title Then a_letter(i) = Capitalize(a_letter(i))
      If onlyFirstLetter Then a_letter(i) = Left(a_letter(i), 1)
    Next
    string = Easp.IIF(hasSpace, Join(a_letter), Join(a_letter, ""))
    ClearPinYin = string
  End Function
  '单词首字母大写
  Public Function Capitalize(ByVal string)
    If Len(string) < 2 Then
      Capitalize = UCase(string)
    Else
      Capitalize = UCase(Left(string, 1)) & Mid(string, 2)
    End If
  End Function
  '把带声调的拼音韵母转换为普通字母
  Private Function SwitchTone(ByVal string, ByVal toneToLetter, ByVal hasToneNumber)
    If hasToneNumber Then
      string = string & "0"
      string = Easp.Str.Replace(string, "^(.*)(ā|ō|ē|ī|ū|ǖ)(.*?)0$", "$1$2$31")
      string = Easp.Str.Replace(string, "^(.*)(á|ó|é|í|ú|ǘ)(.*?)0$", "$1$2$32")
      string = Easp.Str.Replace(string, "^(.*)(ǎ|ǒ|ě|ǐ|ǔ|ǚ)(.*?)0$", "$1$2$33")
      string = Easp.Str.Replace(string, "^(.*)(à|ò|è|ì|ù|ǜ)(.*?)0$", "$1$2$34")
    End If
    If toneToLetter Then
      string = Easp.Str.Replace(string, "ā|á|ǎ|à", "a")
      string = Easp.Str.Replace(string, "ō|ó|ǒ|ò", "o")
      string = Easp.Str.Replace(string, "ē|é|ě|è", "e")
      string = Easp.Str.Replace(string, "ī|í|ǐ|ì", "i")
      string = Easp.Str.Replace(string, "ū|ú|ǔ|ù", "u")
      string = Easp.Str.Replace(string, "ǖ|ǘ|ǚ|ǜ|ü", "v")
    End If
    SwitchTone = string
  End Function

  '在线取得结果的方法原型，返回数组 [拼音, 翻译, 分词]
  Private Function GetResults(ByVal chinese)
    Dim http, s_result, a_http, o_dic, a_word, i, j
    If Not Easp.Str.IsSame(s_cn, chinese) Then
      s_cn = chinese
      Set o_dic = Server.CreateObject("Scripting.Dictionary")
      a_result(0) = s_cn
      a_result(1) = s_cn
      a_result(2) = o_dic.Items
      If Easp.Has(s_cn) And Easp.Str.Test(s_cn, "[\u4e00-\u9fa5]") Then
        Set http = Easp.Http.New
          http.SetHeader "Host:translate.google.cn"
          http.SetHeader "Referer:http://translate.google.cn/"
          http.SetHeader "User-Agent:Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.101 Safari/537.36"
          s_result = http.Get("http://translate.google.cn/translate_a/t?client=t&sl=zh-CN&tl=en&hl=zh-CN&sc=2&ie=UTF-8&oe=UTF-8&oc=2&otf=1&srcrom=1&ssel=6&tsel=3&pc=1&q=" & Server.URLEncode(s_cn))
          Set http = Nothing
          Set a_http = Easp.Decode(Trim(s_result))
          Dim a_http_0, a_http_5
          Set a_http_0 = a_http(0)
          Set a_http_5 = a_http(5)
          a_result(0) = ""
          a_result(1) = ""
          For i = 0 To a_http_0.Length - 1
            a_result(0) = a_result(0) & a_http_0(i)(3)
            a_result(1) = a_result(1) & a_http_0(i)(0)
          Next
          Set a_http_0 = Nothing
          For i = 0 To a_http_5.Length - 1
            a_word = Split(a_http_5(i)(0), " ")
            For j = 0 To Ubound(a_word)
              'Easp.Println a_word(j)
              If Len(a_word(j))>1 And Easp.Str.Test(a_word(j), "[a-zA-Z\u4e00-\u9fa5]") Then
                o_dic.Add "key" & o_dic.Count, a_word(j)
                'Easp.Println "-----------------" & a_word(j)
              End If
            Next
          Next
          Set a_http_5 = Nothing
          Set a_http = Nothing
          a_result(2) = o_dic.Items
      End If
      Set o_dic = Nothing
    End If
    GetResults = a_result
    'Easp.PrintlnString a_result
  End Function  
End Class
%>