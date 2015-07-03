<%
'######################################################################
'## easp.tar.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP File Archiver Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2015-06-01 1:24:16
'## Description :   Pack multiple folders and files into one file or 
'##                 unpack a zipped file to the server.
'######################################################################

Class EasyASP_Tar

  Private BLOCKSIZE
  Private dic_files, o_fso
  Private s_savePath, s_insidePath, s_tarPath, s_downFileName
  Private b_hasSelf

  Private Sub Class_Initialize()
    BLOCKSIZE = 512
    Set dic_files = Server.CreateObject("Scripting.Dictionary")
    Set o_fso = Server.CreateObject("Scripting.FileSystemObject")
    s_savePath = "."
    s_insidePath = ""
    s_tarPath = ""
    b_hasSelf = True
    s_downFileName = ""
    Easp.Error("error-tar-packfaild") = Easp.Lang("error-tar-packfaild")
    Easp.Error("error-tar-unpackfaild") = Easp.Lang("error-tar-unpackfaild")
  End Sub
  
  Private Sub Class_Terminate()
    Set o_fso = Nothing
    Set dic_files = Nothing
  End Sub
  '保存路径
  Public Property Let SavePath(ByVal s_path)
    s_savePath = s_path
  End Property
  Public Property Get SavePath()
    SavePath = s_savePath
  End Property
  '添加文件夹时是否包含自身
  Public Property Let HasSelf(ByVal b)
    b_hasSelf = Easp.IIF(b, True, False)
  End Property
  Public Property Get HasSelf()
    HasSelf = b_hasSelf
  End Property
  '添加文件夹或文件到指定目录
  Public Sub AddTo(ByVal s_path, ByVal s_insidePath)
    Dim s_fName, s_time
    s_insidePath = CreateFolder(s_insidePath)
    s_path = Easp.Fso.MapPath(s_path)
    s_fName = Mid(s_path, InStrRev(s_path, "\")+1)
    'Easp.Console "Add Folder: " & s_path
    If Easp.Fso.isFolder(s_path) Then
      '如果是文件夹
      LoadFolder s_path, s_insidePath & Easp.IfThen(b_hasSelf, s_fName & "/")
    ElseIf Easp.Fso.IsFile(s_path) Then
      '如果是文件
      dic_files(s_insidePath & s_fName) = Array(True, True, s_path, GetFileTimeStamp(s_path))
    Else
      If Easp.Debug Then
        Easp.Error.FunctionName = "Easp.Tar.AddTo"
        Easp.Error.Detail = s_path
        Easp.Error.Raise "error-fso-filenotfound"
      End If
    End If
  End Sub
  '递归添加文件夹内文件
  Private Sub LoadFolder(ByVal s_path, ByVal s_insidePath)
    Dim a_folder, s_file, s_time, s_key, i
    a_folder = Easp.Fso.Dir(s_path) '读取文件夹
    For i = 0 To Ubound(a_folder, 2)
      s_file = s_path & "\" & a_folder(0, i)
      s_time = Easp.Date.ToUnixTimeCn(a_folder(2, i))
      s_key = s_insidePath & a_folder(0, i)
      If Right(s_file, 1) = "/" Then
        dic_files(s_key) = Array(False, True, s_time)
        LoadFolder Left(s_file, Len(s_file)-1), s_key
      Else
        dic_files(s_key) = Array(True, True, s_file, s_time)
      End If
      If Not Response.IsClientConnected Then
        Exit For
      End If
    Next
  End Sub
  '获取文件的最后修改时间
  Private Function GetFileTimeStamp(ByVal s_path)
    Dim f
    Set f = o_fso.GetFile(s_path)
    GetFileTimeStamp = Easp.Date.ToUnixTimeCn(f.DateLastModified)
    Set f = Nothing
  End Function
  '添加文件夹或文件
  Public Sub Add(ByVal s_path)
    AddTo s_path, ""
  End Sub
  '创建文件夹
  Public Function CreateFolder(ByVal s_path)
    Dim arr, i, tmp : tmp = ""
    s_path = Replace(s_path, "\", "/")
    arr = Split(s_path, "/")
    For i = 0 To Ubound(arr)
      If Easp.Has(arr(i)) Then
        tmp = tmp & arr(i) & "/"
        If Not dic_files.Exists(tmp) Then
          dic_files(tmp) = Array(False, False, Easp.Date.GetTimeStamp())
        End If
      End If
    Next
    CreateFolder = tmp
  End Function
  '添加文本
  Public Sub CreateFile(ByVal s_fileName, ByVal s_content)
    Dim s_path, s_fName
    If Instr(s_fileName, "/") > 0 Then
      s_path = Left(s_fileName, InstrRev(s_fileName, "/")-1)
      s_path = CreateFolder(s_path)
      s_fName = Mid(s_fileName, InStrRev(s_fileName, "/")+1)
      s_fileName = s_path & s_fName
    End If
    dic_files(s_fileName) = Array(True, False, s_content, Easp.Date.GetTimeStamp())
  End Sub
  '打包
  Public Function Pack()
    Pack = PackTo(s_savePath)
  End Function
  '打包保存到
  Public Function PackTo(ByVal s_filePath)
    On Error Resume Next
    Dim p, o_tar, key, arr, o_down
    If Easp.Has(s_filePath) Then
      s_savePath = s_filePath
      p = Easp.Fso.MapPath(s_filePath)
      Call Easp.Fso.MD(Left(p,InstrRev(p,"\")-1))
    End If
    Set o_tar = Server.CreateObject("ADODB.Stream")
    With o_tar
      .Type = 2
      .Charset = "x-ansi"
      .Open
      .Position = 0
      For Each key In dic_files
        arr = dic_files(key)
        If arr(0) Then '文件
          If arr(1) Then '真实文件
            AddToTar o_tar, key, LoadFile(arr(2)), arr(3)
          Else '待生成的文件
            AddToTar o_tar, key, LoadText(arr(2), Null), arr(3)
          End If
        Else '文件夹
          AddToTar o_tar, key, Null, arr(2)
        End If
        If Not Response.IsClientConnected Then
          Exit For
        End If
      Next
      If Not Response.IsClientConnected Then PackTo = False : Exit Function
      .WriteText String(1024, Chr(0))
      If Easp.Has(s_filePath) Then
        .SaveToFile p, 2
      Else
        OutPutFile o_tar, Easp.IfHas(s_downFileName, "file_" & Easp.Date.Format(Now, "ymmddhhiiss") & ".tar")
      End If
      .Close
    End With
    Set o_tar = Nothing
    If Err.Number <> 0 Then
      PackTo = False
      If Easp.Debug Then
        Easp.Error.FunctionName = "Easp.Tar.PackTo"
        Easp.Error.Detail = p
        Easp.Error.Raise "error-tar-packfaild"
      End If
    Else
      PackTo = True
    End If
  End Function
  '直接下载打包文件
  Public Sub DownLoad(ByVal s_name)
    If InStr(s_name, "/") > 0 Or Instr(s_name, "\") > 0 Then
      s_name = Replace(s_name, "\", "/")
      s_name = Mid(s_name, InStrRev(s_name, "/")+1)
    End If
    s_downFileName = s_name
    PackTo Null
  End Sub
  '添加内容到文件流
  Private Sub AddToTar(ByRef o_tar, ByVal s_pathname, ByRef o_file, ByVal timeStamp)
    Dim pos, is_folder, o_filename, checkNum, i, sizeMod, txt
    is_folder = Right(s_pathname, 1) = "/"
    With o_tar
      pos = .Position
      '文件名
      Set o_filename = LoadText(s_pathname & String(100-Easp.Str.Leng(s_pathname), Chr(0)), "gbk")
      o_filename.CopyTo o_tar
      Set o_filename = Nothing
      '类型
      .WriteText Easp.IIF(is_folder, " 4", "10")
      .WriteText "0777 " & Chr(0)
      .WriteText "     0 " & Chr(0)
      .WriteText "     0 " & Chr(0)
      '大小
      If is_folder Then
        .WriteText "          0 "
      Else
        If TypeName(o_file) = "Stream" Then
          .WriteText Right("           " & Oct(o_file.Size), 11) & " "
        Else
          .WriteText "          0 "
        End If
      End If
      '最后修改时间
      .WriteText Right("00000000000" & Oct(timeStamp), 11) & Chr(0)
      '校验和留空
      .WriteText "       " & Chr(0)
      '链接
      .WriteText Easp.IIF(is_folder, "5", "0")
      .WriteText String(355, Chr(0))
      '写入文件
      If TypeName(o_file) = "Stream" Then
        o_file.CopyTo o_tar
        sizeMod = (o_file.Size Mod BLOCKSIZE)
        If sizeMod > 0 Then
          .WriteText String(BLOCKSIZE - sizeMod, Chr(0))
        End If
      End If
      '校验和
      checkNum = 32
      .Position = pos
      For i = 1 To BLOCKSIZE
        txt = .ReadText(1)
        checkNum = checkNum + (AscB(txt) And &HFF&)
      Next
      .Position = pos + 148
      .WriteText Right(String(6," ") & Oct(checkNum),6) & " "
      .Position = .Size
    End With
  End Sub
  '载入tar文件
  Public Sub LoadTar(ByVal s_path)
    s_tarPath = s_path
  End Sub
  '解包tar
  Public Function UnPack()
    UnPack = UnPackTo(s_tarPath, s_savePath)
  End Function
  '解包tar到指定目录
  Public Function UnPackTo(ByVal s_tarFilePath, ByVal s_folderPath)
    On Error Resume Next
    Dim o_tar, o_data
    Dim s_header, s_fileName, s_type, i_fileSize, is_file
    If Easp.IsN(s_folderPath) Then s_folderPath = "."
    s_savePath = s_folderPath
    s_folderPath = Easp.Fso.MapPath(s_folderPath)
    If b_hasSelf Then
      s_folderPath = s_folderPath & "\" & Easp.Fso.NameOf(s_tarFilePath)
      s_savePath = s_savePath & Easp.IfThen(Right(s_savePath, 1) <> "/", "/") & Easp.Fso.NameOf(s_tarFilePath)
    End If
    Easp.Fso.MD s_folderPath
    Set o_data = Server.CreateObject("ADODB.Stream")
    o_data.Type = 1 : o_data.Open
    Set o_file = Server.CreateObject("ADODB.Stream")
    o_file.Type = 1 : o_file.Open
    Set o_tar = LoadFile(s_tarFilePath)
    With o_tar
      .CopyTo o_data
      .Position = 0
      o_data.Position = 0
      Do While(True)
        is_file = False
        i_fileSize = 0
        s_fileName = .ReadText(100) '读文件名
        s_fileName = Left(s_fileName, Instr(s_fileName, Chr(0)) - 1)
        s_fileName = ChangeCharset(s_fileName, "x-ansi", "gbk")
        s_fileName = Replace(s_fileName, "/", "\")
        s_type = .ReadText(3) '读类型
        If Right(s_fileName, 1) = "\" And Right(s_type, 2) = "40" Then
          '如果是文件夹
          Easp.Fso.MD s_folderPath & "\" & Left(s_fileName, Len(s_fileName)-1)
          'Easp.Console s_folderPath & "\" & Left(s_fileName, Len(s_fileName)-1)
          .ReadText(32)
        Else
          '如果是文件，取出文件大小
          i_fileSize = Int("&0" & Int(Right(.ReadText(32), 11)))
          'Easp.Console s_folderPath & "\" & s_fileName & "|" & i_fileSize
          is_file = True
        End If
        '跳过剩余的文件头
        .ReadText(377)
        o_data.Read(BLOCKSIZE)
        If is_file Then
          '如果是文件，就将文件内容提取出来并保存
          If i_fileSize > 0 Then
            Easp.Fso.SaveAs s_folderPath & "\" & s_fileName, o_data.Read(i_fileSize)
            .ReadText(i_fileSize)
          Else
            Easp.Fso.CreateFile s_folderPath & "\" & s_fileName, ""
          End If
        End If
        '跳过所有的填充块
        Do While (.ReadText(1)=Chr(0))
        Loop
        If o_tar.Eos Then Exit Do
        .Position = .Position - 1
        o_data.Position = .Position
        If Not Response.IsClientConnected Then
          Exit Do
        End If
      Loop
    End With
    If Not Response.IsClientConnected Then UnPackTo = False : Exit Function
    o_tar.Close : Set o_tar = Nothing
    o_data.Close : Set o_data = Nothing
    If Err.Number <> 0 Then
      UnPackTo = False
      If Easp.Debug Then
        Easp.Error.FunctionName = "Easp.Tar.UnPackTo"
        Easp.Error.Detail = p
        Easp.Error.Raise "error-tar-unpackfaild"
      End If
    Else
      UnPackTo = True
    End If
  End Function
  '列出已加入的文件
  Public Function List()
    Dim key
    Set List = Easp.Json.NewArray
    For Each key In dic_files
      List.Add key
    Next
  End Function
  '读文件到流
  Private Function LoadFile(ByVal s_path)
    On Error Resume Next
    Set LoadFile = Server.CreateObject("ADODB.Stream")
    LoadFile.Type = 2
    LoadFile.CharSet = "x-ansi"
    LoadFile.Open
    LoadFile.LoadFromFile Easp.Fso.MapPath(s_path)
    If Err.Number <> 0 Then
      If Easp.Debug Then
        Easp.Error.FunctionName = "Easp.Tar.UnPackTo"
        Easp.Error.Detail = s_path
        Easp.Error.Raise "error-fso-filenotfound"
      End If
    End If
  End Function
  '写文本到流
  Private Function LoadText(ByVal s_content, ByVal charset)
    If Easp.IsN(charset) Then charset = Easp.CharSet
    Set LoadText = Server.CreateObject("ADODB.Stream")
    LoadText.Type = 2
    LoadText.CharSet = charset
    LoadText.Open
    LoadText.WriteText s_content
    LoadText.Position = 0
    LoadText.CharSet = "x-ansi"
  End Function
  '更改字符串编码
  Private Function ChangeCharset(s, oldChar, newChar)
    Dim o_strm
    Set o_strm = Server.CreateObject("ADODB.Stream")
    With o_strm
      .Type = 2
      .Mode = 3
      .Open
      .CharSet = oldChar
      .WriteText s
      .Position = 0
      .Type = 2
      .CharSet = newChar
      ChangeCharset = .ReadText
      .Close
    End With
    Set o_strm = Nothing
  End Function
  '把流输出到浏览器
  Private Function OutPutFile(o_strm, s_fileName)
    Dim char, sent
    sent = 0
    OutPutFile = True
    o_strm.Position = 0
    o_strm.Type = 1
    Response.AddHeader "content-type", "application/octec-stream"
    Response.AddHeader "Content-Disposition","attachment;filename=" & s_fileName 
    Response.AddHeader "content-length", o_strm.Size
    Do While Not o_strm.EOS
      char = o_strm.Read(1)
      Response.BinaryWrite(char)
      sent = sent + 1
      If (sent MOD 16384) = 0 Then
        Response.Flush
        If Not Response.IsClientConnected Then
          OutPutFile = False
          Exit Do
        End If
      End If
    Loop
    Response.Flush
    if Not Response.IsClientConnected Then OutPutFile = False
  End Function
End Class
%>