<%
'######################################################################
'## easp.upload.asp
'## -------------------------------------------------------------------
'## Feature   :  EasyASP Upload Files Class
'## Version   :  MoLibUpload V1.1
'## Author   :  Anlige(zhanghuiguoanlige@126.com, http://dev.mo.cn)
'## Update Date :  2015-05-10 10:16:49
'## Description :  Upload files with a post form.
'##
'######################################################################
Class EasyASP_MoLibUpload
  Private Form, Fils,StreamT,mvarClsName, mvarClsDescription,mvarSavePath,mvarCheckImageFormat
  Private vCharSet, vMaxSize, vSingleSize, vErr, vVersion, vTotalSize, vExe, vErrExe,vboundary, vLostTime, vFileCount,StreamOpened
  private vMuti,vServerVersion,mvarDescription
  Public IsUploaded, FormArray, s_errLang 'added by EasyASP
  '设置允许上传的总大小
  Public Property Let AllowMaxSize(ByVal value)
    vMaxSize = value
  End Property
  '设置允许上传的单个文件大小
  Public Property Let AllowMaxFileSize(ByVal value)
    vSingleSize = value
  End Property
  '设置允许上传的文件类型
  Public Property Let AllowFileTypes(ByVal value)
    vExe = LCase(value)
    vExe = replace(vExe,"*.","")
    vExe = replace(vExe,";","|")
  End Property
  '检测图片文件格式
  Public Property Let CheckImageFormat(byval value)
    mvarCheckImageFormat = value
  End Property
  '设置文件编码
  Public Property Let CharSet(ByVal value)
    vCharSet = value
  End Property
  '设置上传文件保存路径
  Public Property Let SavePath(ByVal value)
    mvarSavePath = value
  End Property
  '获取上传文件个数
  Public Property Get FileCount()
    FileCount = Fils.count
  End Property
  
  Public Property Get Description()
    Description = mvarDescription
  End Property
  
  Public Property Get Version()
    Version = vVersion
  End Property
  '获取上传文件的总大小
  Public Property Get TotalSize()
    TotalSize = vTotalSize
  End Property
  '获取上传使用的时间，不包括保存文件的时间
  Public Property Get LostTime()
    LostTime = vLostTime
  End Property
  '读取和设置错误提示的语言类型
  Public Property Let ErrorLang(ByVal string)
    s_errLang = string
  End Property
  Public Property Get ErrorLang()
    ErrorLang = s_errLang
  End Property
  
  Private Sub Class_Initialize()
    Dim T__
    Set Form = Server.CreateObject("Scripting.Dictionary")
    Set FormArray = Server.CreateObject("Scripting.Dictionary") 'added by EasyASP
    Set Fils = Server.CreateObject("Scripting.Dictionary")
    Set StreamT = Server.CreateObject("Adodb.stream")
    s_errLang = "en"
    vVersion = "MoLibUpload V1.1"
    vMaxSize = -1
    vSingleSize = -1
    vErr = -1
    vExe = ""
    vTotalSize = 0
    vCharSet = "utf-8"
    StreamOpened=false
    vMuti="_" & Getname() & "_"
    mvarCheckImageFormat = false
    vServerVersion = 6.0
    T__ = lcase(Request.ServerVariables("SERVER_SOFTWARE"))
    T__ = replace(T__,"microsoft-iis/","")
    if isnumeric(T__) then vServerVersion = cdbl(T__)
    mvarClsName = "MoLibFileExtern_"+Getname()
    mvarClsDescription="Class {ClsName}\nPublic ContentType,Size,UserSetName,Path,Position,FormName,TempFormName, NewName,FileName,LocalName,IsFile,Extend,Succeed,Exception,Width,Height,IsImage\nEnd Class"
    mvarClsDescription = Replace(mvarClsDescription,"{ClsName}",mvarClsName)
    mvarClsDescription = Replace(mvarClsDescription,"\n",vbcrlf)
    ExecuteGlobal mvarClsDescription
  End Sub
  
  Private Sub Class_Terminate()
    Dim f
    Form.RemoveAll()
    For each f in Fils
      Set Fils(f) = Nothing
    Next
    Fils.RemoveAll()
    Set FormArray = Nothing 'added by EasyASP
    Set Form = Nothing
    Set Fils = Nothing
    if StreamOpened then StreamT.close()
    Set StreamT = Nothing
  End Sub
  
  Private Function ParseSizeLimit(byval SizeLimit)
    dim unit,value,multiplier,limit
    If Not isnumeric(SizeLimit) Then
      multiplier = 1
      SizeLimit = ReplaceEx(lcase(SizeLimit),"\s","")
      value = replaceex(SizeLimit,"[^\d]+","")
      if isnumeric(value) then
        value = clng(value)
        if right(SizeLimit,2)="gb" then multiplier = 1073741824
        if right(SizeLimit,2)="mb" then multiplier = 1048576
        if right(SizeLimit,2)="kb" then multiplier = 1024
        limit = value * multiplier
      else
        limit=-1
      end if
    else
      limit = SizeLimit
    End If  
    if limit<-1 then limit=-1
    ParseSizeLimit = limit
  End Function
  '开始上传动作
  Public Function GetData()
    Dim oarr 'added by EasyASP
    GetData =false
    vMaxSize = ParseSizeLimit(vMaxSize)
    vSingleSize = ParseSizeLimit(vSingleSize)
    Dim time1
    time1 = timer()
    Dim value, str, bcrlf, fpos, sSplit, slen, istart,ef
    Dim TotalBytes,tempdata,BytesRead,ChunkReadSize,PartSize,DataPart,formend, formhead, startpos, endpos, formname, FileName, fileExe, valueend, NewName,localname,type_1,contentType
    TotalBytes = Request.TotalBytes
    ef = false
    If checkEntryType = false Then ef = true : mvarDescription = Easp.Lang("error-uplaod-enctypeor-" & s_errLang)
    If vServerVersion>=6 Then
      If Not ef Then
        If vMaxSize > 0 And TotalBytes > vMaxSize Then ef = true : mvarDescription = Easp.Lang("error-uplaod-filemaxsize-" & s_errLang)
      End If
    End If
    If ef Then Exit function
    vTotalSize = 0 
    StreamT.Type = 1
    StreamT.Mode = 3
    StreamT.Open
    StreamOpened = true
    BytesRead = 0
    ChunkReadSize = 1024 * 16
    Do While BytesRead < TotalBytes
      PartSize = ChunkReadSize
      If PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
      DataPart = Request.BinaryRead(PartSize)
      StreamT.Write DataPart
      BytesRead = BytesRead + PartSize
    Loop
    StreamT.Position = 0
    tempdata = StreamT.Read
    bcrlf = ChrB(13) & ChrB(10)
    fpos = InStrB(1, tempdata, bcrlf)
    sSplit = MidB(tempdata, 1, fpos - 1)
    slen = LenB(sSplit)
    istart = slen + 2
    Do
      formend = InStrB(istart, tempdata, bcrlf & bcrlf)
      if formend<=0 then exit do
      formhead = MidB(tempdata, istart, formend - istart)
      str = Bytes2Str(formhead)
      startpos = InStr(str, "name=""") + 6
      if startpos<=0 then exit do
      endpos = InStr(startpos, str, """")
      if endpos<=0 then exit do
      formname = LCase(Mid(str, startpos, endpos - startpos))
      valueend = InStrB(formend + 3, tempdata, sSplit)
      if valueend<=0 then exit do
      If InStr(str, "filename=""") > 0 Then
        formname = formname & vMuti & "0"
        startpos = InStr(str, "filename=""") + 10
        endpos = InStr(startpos, str, """")
        type_1=instr(endpos,lcase(str),"content-type")
        contentType=lcase(trim(mid(str,type_1+13)))
        FileName = Mid(str, startpos, endpos - startpos)
        If Trim(FileName) <> "" Then
          FileName = Replace(FileName, "/", "\")
          FileName = Replace(FileName, chr(0), "")
          LocalName = FileName
          FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
          If instr(FileName,".")>0 Then
            fileExe = Split(FileName, ".")(UBound(Split(FileName, ".")))
          else
            fileExe = ""
          End If
          If vExe <> "" Then
            If checkExe(fileExe) = True Then
              mvarDescription = Easp.Lang("error-uplaod-filetype-" & s_errLang) & "(." & ucase(fileExe) & ")"
              vErrExe = fileExe
              tempdata = empty
              Exit function
            End If
          End If
          NewName = Getname()
          vTotalSize = vTotalSize + valueend - formend - 6
          If vSingleSize > 0 And (valueend - formend - 6) > vSingleSize Then
            mvarDescription = Easp.Lang("error-uplaod-filesize-" & s_errLang)
            tempdata = empty
            Exit function
          End If
          If vMaxSize > 0 And vTotalSize > vMaxSize Then
            mvarDescription = Easp.Lang("error-uplaod-filemaxsize-" & s_errLang)
            tempdata = empty
            Exit function
          End If
          If Fils.Exists(formname) Then formname = GetNextFormName(formname)
          fileExe = lcase(fileExe)
          fileExe = replace(fileExe,";","")
          Dim fileCls:set fileCls= NewFile()
          fileCls.ContentType=contentType
          fileCls.Size= (valueend - formend - 6)
          fileCls.Position = (formend + 3)
          fileCls.FormName = mid(formname,instr(formname,vMuti)-1)
          fileCls.TempFormName = formname
          fileCls.NewName = NewName & "." & fileExe
          fileCls.FileName = FileName
          fileCls.LocalName = FileName
          fileCls.IsFile = true
          fileCls.IsImage = false
          fileCls.Extend=fileExe
          if mvarCheckImageFormat=true then
            if Instr(",image/jpeg,image/pjpeg,image/jpg,image/gif,image/png,image/bmp,application/x-shockwave-flash,","," & contentType & ",")>0 or Instr(",jpg,jpeg,pjpeg,bmp,png,gif,swf,","," & fileExe & ",")>0 then
              fileCls.IsImage = IsImage(fileCls)
              if fileCls.IsImage then fileCls.NewName = NewName & "." & fileCls.Extend
            end if
          end if
          Fils.Add formname, fileCls
        End If
      Else
        value = MidB(tempdata, formend + 4, valueend - formend - 6)
        If Form.Exists(formname) Then
          Form(formname) = Form(formname) & ", " & Bytes2Str(value)
          'added by EasyASP start
          Set oarr = FormArray(formname)
          oarr.Add Bytes2Str(value)
          Set FormArray(formname) = oarr
          Set oarr = Nothing
          'added by EasyASP end
        Else
          Form.Add formname, Bytes2Str(value)
          'added by EasyASP start
          Set oarr = Easp.Json.NewArray
          oarr.Add Bytes2Str(value)
          FormArray.Add formname, oarr
          Set oarr = Nothing
          'added by EasyASP end
        End If
      End If
      istart = valueend + 2 + slen
    Loop Until (istart + 2) >= LenB(tempdata)
    tempdata = empty
    vLostTime = FormatNumber((timer-time1)*1000,2)
    'added by EasyASP start
    Dim postKey
    For Each postKey In FormArray
      Easp.Var("post." & postKey) = FormArray(postKey).GetArray()
    Next
    IsUploaded = True
    'added by EasyASP end
    GetData =true
  End Function
  
  Private Function CheckExe(ByVal ex)
    Dim notIn: notIn = True
    If vExe="*" then
      notIn=false 
    elseIf InStr(1, vExe, "|") > 0 Then
      Dim tempExe: tempExe = Split(vExe, "|")
      Dim I: I = 0
      For I = 0 To UBound(tempExe)
        If LCase(ex) = tempExe(I) Then
          notIn = False
          Exit For
        End If
      Next
    Else
      If vExe = LCase(ex) Then
        notIn = False
      End If
    End If
    checkExe = notIn
  End Function
  
  Private Function Bytes2Str(ByVal byt)
    If LenB(byt) = 0 Then
      Bytes2Str = ""
      Exit Function
    End If
    Dim mystream, bstr
    Set mystream =Server.CreateObject("ADODB.Stream")
    mystream.Type = 2
    mystream.Mode = 3
    mystream.Open
    mystream.WriteText byt
    mystream.Position = 0
    mystream.CharSet = vCharSet
    mystream.Position = 2
    bstr = mystream.ReadText()
    mystream.Close
    Set mystream = Nothing
    Bytes2Str = bstr
  End Function
  
  Private Function Getname()
    Dim y, m, d, h, mm, S, r
    Randomize
    y = Year(Now)
    m = right("0" & Month(Now),2)
    d = right("0" & Day(Now),2)
    h = right("0" & Hour(Now),2)
    mm =right("0" & Minute(Now),2)
    S = right("0" & Second(Now),2)
    r = CInt(Rnd() * 10000)
    r = right("0000" & r,4)
    Getname = y & m & d & h & mm & S & r
  End Function
  '检测提交表单类型
  Public Function checkEntryType()
    Dim ContentType, ctArray, bArray,RequestMethod
    RequestMethod=trim(LCase(Request.ServerVariables("REQUEST_METHOD")))
    if RequestMethod="" or RequestMethod<>"post" then
      checkEntryType = False
      exit function
    end if
    ContentType = LCase(Request.ServerVariables("HTTP_CONTENT_TYPE"))
    if ContentType="" then ContentType = LCase(Request.ServerVariables("CONTENT_TYPE"))
    ctArray = Split(ContentType, ";")
    if ubound(ctarray)>=0 then
      If Trim(ctArray(0)) = "multipart/form-data" Then
      checkEntryType = True
      vboundary = Split(ContentType,"boundary=")(1)
      Else
      checkEntryType = False
      End If
    else
      checkEntryType = False
    end if
  End Function
  '获取表单数据
  Public Function Post(ByVal formname)
    If trim(formname) = "-1" Then
      Set Post = Form
    Else
      If Form.Exists(LCase(formname)) Then
        Post = Form(LCase(formname))
      Else
        Post = ""
      End If
    End If
  End Function
  '获取上传后的文件
  Public Default Function Files(ByVal formname)
    If trim(formname) = "-1" Then
      Set Files = Fils
    Else
      dim vname
      vname = LCase(formname) & vMuti & "0"
      if instr(formname,vMuti)>0 then vname = formname
      If Fils.Exists(vname) Then
        Set Files = Fils(vname)
      Else
        Set Files = NewFile()
        Files.IsFile = false
      End If
    End If
  End Function
  '筛选上传后的文件
  Public Function Search(ByVal formname)
    if formname="*" or formname="-1" then
      Set Search = Fils
      Exit Function
    end if
    Dim TempFormName
    TempFormName = formname & vMuti
    Dim FileCollection
    Set FileCollection = Server.CreateObject("Scripting.Dictionary")
    Dim v
    For Each v In Fils
      If lcase(left(v,len(TempFormName))) = lcase(TempFormName) Then
        FileCollection.Add v,Fils(v)
      End If
    Next
    Set Search = FileCollection
  End Function
  '快速保存指定文件域的文件
  Public Function QuickSave(ByVal formname)
    Dim FC,SucceedCount,File
    SucceedCount = 0
    Set FC = Search(formname)
    For Each File In FC
      If Save(File,0,True).Succeed Then SucceedCount = SucceedCount + 1
    Next
    QuickSave = SucceedCount
  End Function
  '判断是否是合法的图像文件
  Public Function IsImage(Name)
    Dim File
    if not isobject(Name) then
      Set File = Files(Name)
      If Not File.IsFile Then
        IsImage = false
        Exit Function
      End If
    else
      Set File = Name
    end if
    IsImage = false
    'from internet
    Dim intTemp,strTemp,isJpeg,ispng,a,b,c,d,tempContentType
    tempContentType = lcase(File.contentType)
    select case lcase(File.Extend)
      case "jpg","jpeg","pjpeg"
        tempContentType = "image/jpeg"
      case "gif"
        tempContentType = "image/gif"
      case "png"
        tempContentType = "image/png"
      case "bmp"
        tempContentType = "image/bmp"
      case "swf"
        tempContentType = "application/x-shockwave-flash"
    end select
    select case tempContentType
    case "image/jpeg","image/pjpeg","image/jpg"
      if Lcase(File.Extend)<>"jpg" then File.Extend="jpg"
      StreamT.Position=File.Position+3
      do while not StreamT.EOS
        do
          intTemp = Ascb(StreamT.Read(1))
        loop while intTemp = 255 and not StreamT.EOS
        if intTemp < 192 or intTemp > 195 then
          StreamT.read(Bin2Val(StreamT.Read(2))-2)
        else
          Exit do
        end if
        do
          intTemp = Ascb(StreamT.Read(1))
        loop while intTemp < 255 and not StreamT.EOS
      loop
      StreamT.Read(3)
      File.Height = Bin2Val(StreamT.Read(2))
      File.Width = Bin2Val(StreamT.Read(2))
      StreamT.Position = File.Position + File.Size-2
      isJpeg = false
      if ascb(StreamT.read(1))=&HFF then
        if ascb(StreamT.read(1))=&HD9 then
          isJpeg = true
        end if
      end if
      if not isJpeg then
        File.Width = 0
        File.Height = 0
        File.size=0
      end if
    case "image/gif"
      if Lcase(File.Extend)<>"gif" then File.Extend="gif"
      StreamT.Position=File.Position+6
      File.Width = BinVal2(StreamT.Read(2))
      File.Height = BinVal2(StreamT.Read(2))
      StreamT.Position = File.Position + File.Size-1
      if ascb(StreamT.read(1)) <>asc(";") then
        File.Width = 0
        File.Height = 0
        File.size=0
      end if
    case "image/png"
      if Lcase(File.Extend)<>"png" then File.Extend="png"
      StreamT.Position=File.Position+18
      File.Width = Bin2Val(StreamT.Read(2))
      StreamT.Read(2)
      File.Height = Bin2Val(StreamT.Read(2))
      StreamT.Position = File.Position + File.Size-12
      ispng = false
      intTemp = Ascb(StreamT.Read(1))
      intTemp = intTemp+Ascb(StreamT.Read(1))
      intTemp = intTemp+Ascb(StreamT.Read(1))
      intTemp = intTemp+Ascb(StreamT.Read(1))
      if intTemp=0 then
        a = Ascb(StreamT.Read(1))
        b = Ascb(StreamT.Read(1))
        c = Ascb(StreamT.Read(1))
        d = Ascb(StreamT.Read(1))
        if a=&H49 and b = &H45 and c = &H4e and d = &H44 then
          a = Ascb(StreamT.Read(1))
          b = Ascb(StreamT.Read(1))
          c = Ascb(StreamT.Read(1))
          d = Ascb(StreamT.Read(1))
          if a=&HAE and b = &H42 and c = &H60 and d = &H82 then
            ispng = true  
          end if
        end if
      end if
      if not ispng then
        File.Width = 0
        File.Height = 0
        File.size=0
      end if
    case "image/bmp"
      if Lcase(File.Extend)<>"bmp" then File.Extend="bmp"
      StreamT.Position=File.Position+18
      File.Width = BinVal2(StreamT.Read(4))
      File.Height = BinVal2(StreamT.Read(4))
      StreamT.Position=File.Position+2
      File.Size= BinVal2(StreamT.Read(4))
    case "application/x-shockwave-flash"
      if Lcase(File.Extend)<>"swf" then File.Extend="swf"
      StreamT.Position=File.Position
      if Ascb(StreamT.Read(1))=&H46 then
        StreamT.Position=File.Position+8
        strTemp = Num2Str(Ascb(StreamT.Read(1)), 2 ,8)
        intTemp = Str2Num(Left(strTemp, 5), 2)
        strTemp = Mid(strTemp, 6)
        while (Len(strTemp) < intTemp * 4)
          strTemp = strTemp & Num2Str(Ascb(StreamT.Read(1)), 2 ,8)
        wend
        File.Width = Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
        File.Height = Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
      end if
    end select
    if File.Width>0 and File.Height>0 then
      IsImage = true
    else
      IsImage = false
      File.Size=0
    end if
  End Function
  '保存文件
  Public Function Save(Byref Name,byval tOption, byval OverWrite)
    Dim File
    if not isobject(name) then
      Set File = Files(Name)
      If Not File.IsFile Then
        File.Succeed = false
        File.Exception = Easp.Lang("error-uplaod-fileno-" & s_errLang)
        Set Save = File
        Exit Function
      End If
    else
      Set File = Name
    end if
    If Not File.IsFile Then
      File.Succeed = false
      File.Exception = Easp.Lang("error-uplaod-fileno-" & s_errLang)
      Set Save = File
      Exit Function
    End If
    On Error Resume Next
    Err.clear
    Dim IsP,Path
    Path = mvarSavePath
    IsP = (InStr(mvarSavePath, ":") = 2)
    If Not IsP Then Path = Server.MapPath(mvarSavePath)
    Path = Replace(Path, "/", "\")
    If Mid(Path, Len(Path) - 1) <> "\" Then Path = Path + "\"
    CreateFolder Path
    File.Path= Replace(Replace(Path,Server.MapPath("/"),""),"\","/")
    If tOption = 1 Then
      Path = Path & File.LocalName: File.FileName =File.LocalName
    Else
      If tOption = -1 And File.UserSetName <> "" Then
        Path = Path & File.UserSetName & "." & File.Extend: File.FileName = File.UserSetName & "." & File.Extend
      Else
        Path = Path & File.NewName: File.FileName = File.NewName
      End If
    End If
    If Not OverWrite Then
      Path = GetFilePath(File)
    End If
    If Err.Number<>0 Then
      File.Succeed = false
      File.Exception=Err.Description
      Err.clear()
      Set Save = File
      Exit Function
    End if
    Dim tmpStrm
    Set tmpStrm =Server.CreateObject("ADODB.Stream")
    tmpStrm.Mode = 3
    tmpStrm.Type = 1
    tmpStrm.Open
    StreamT.Position = File.Position
    StreamT.copyto tmpStrm,File.Size
    tmpStrm.SaveToFile Path, 2
    tmpStrm.Close
    Set tmpStrm = Nothing
    If Err.Number=0 Then
      File.Succeed = true
    Else
      File.Succeed = false
      File.Exception=Err.Description
      Err.clear()
    End If
    Set Save = File
  End Function
  '获取已上传文件的二进制流
  Public Function GetBinary(byval Name)
    Dim File
    Set File = Files(Name)
    If Not File.IsFile Then
      GetBinary = chrb(0)
      Exit Function
    End If
    StreamT.Position = File.Position
    GetBinary = StreamT.read(File.Size)
  End Function 
  
  Private Function GetNextFormName(byval formname)
    Dim formStart,currentIndex
    formStart = left(formname,instr(formname,vMuti)+len(vMuti)-1)
    currentIndex = mid(formname,instr(formname,vMuti)+len(vMuti))
    currentIndex =cint(currentIndex)
    do while Fils.Exists(formname)
      currentIndex = currentIndex + 1
      formname = formStart & currentIndex
    loop
    GetNextFormName = formname
  End Function
  Private Function ReplaceEx(sourcestr, regString, str)
    if isnull(sourcestr) then sourcestr=""
    dim re
    Set re = new RegExp
    re.IgnoreCase = true
    re.Global = True
    re.pattern = "" & regString & ""
    str = re.replace(sourcestr, str)
    set re = Nothing
    ReplaceEx = str
  End Function
  Private Function CreateFolder(ByVal folderPath )
    Dim oFSO
    Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    Dim sParent 
    sParent = oFSO.GetParentFolderName(folderPath)
    If sParent = "" Then Exit Function
    If Not oFSO.FolderExists(sParent) Then CreateFolder (sParent)
    If Not oFSO.FolderExists(folderPath) Then oFSO.CreateFolder (folderPath)
    Set oFSO = Nothing
  End Function
  
  Private Function GetFilePath(Byref File) 
    Dim oFSO, Fname , FNameL , i 
    i = 0
    Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    Fname = Server.MapPath(File.Path & File.FileName)
    FNameL = Mid(File.FileName, 1, InStr(File.FileName, ".") - 1)
    Do While oFSO.FileExists(Fname)
      Fname = Server.MapPath(File.Path & FNameL & "(" & i & ")." & File.Extend)
      File.FileName = FNameL & "(" & i & ")." & File.Extend
      i = i + 1
    Loop
    Set oFSO = Nothing
    GetFilePath = Fname
  End Function
  
  Private Function NewFile()
    Execute "Set NewFile = new " & mvarClsName
    NewFile.Width = 0
    NewFile.Height = 0
  End Function
  
  Private Function BinVal2(bin)
    dim lngValue,i
    lngValue=0
    for i = lenb(bin) to 1 step -1
      lngValue = lngValue *256 + Ascb(midb(bin,i,1))
    Next
    BinVal2=lngValue
  End Function

  Private Function Bin2Val(bin)
    dim lngValue,i
    lngValue=0
    for i = 1 to lenb(bin)
      lngValue = lngValue *256 + Ascb(midb(bin,i,1))
    Next
    Bin2Val=lngValue
  End Function

  Private Function Num2Str(num, base, lens)
    Dim ret,i
    ret = ""
    while(num >= base)
      i  = num Mod base
      ret = i & ret
      num = (num - i) / base
    wend
    Num2Str = Right(String(lens, "0") & num & ret, lens)
  End Function

  Private Function Str2Num(str, base)
    Dim ret, i
    ret = 0 
    for i = 1 to Len(str)
      ret = ret * base + Cint(Mid(str, i, 1))
    Next
    Str2Num = ret
  End Function
End Class
%>