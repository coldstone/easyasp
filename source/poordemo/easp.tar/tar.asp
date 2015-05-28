<!--#include file="../../easyasp/easp.asp" --><%
Dim act : act = Easp.Var("act")
Select Case act
  Case "pack", "down"
    Dim i
    '在压缩包内创建新文件
    Easp.Tar.CreateFile Easp.Var("newfile"), Easp.Var("newfilecontent")
    '在压缩包内创建新文件夹
    Easp.Tar.CreateFolder Easp.Var("newfolder")
    For i = 0 To Ubound(Easp.Var("files_array"))
      '添加文件夹或文件到压缩包
      If Left(Easp.Var("files_array_" & i), 1) <> "|" Then
        Easp.Tar.Add Easp.Var("files_array_" & i)
      End If
    Next
    If act = "pack" Then
      '打包保存到硬盘
      If Easp.Tar.PackTo(Easp.Var("tarfile")) Then
        Easp.Str.JsAlertUrl "打包成功！文件存放到站点的下面位置：" & vbCrLf & Easp.Tar.SavePath, "."
      Else
        Easp.Str.JsAlertUrl "打包失败", "."
      End If
    Else
      '打包后直接输出到浏览器让用户下载
      Easp.Tar.DownLoad Easp.Var("tarfile")
      Easp.RR "."
    End If
  Case "unpack"
    Easp.Tar.HasSelf = Easp.Var("hasself") = "1"
    '方法一：
    'Easp.Tar.SavePath = Easp.Var("savepath")
    'Easp.Tar.LoadTar Easp.Var("untarfile")
    'Easp.Tar.Unpack()
    '方法二：
    Easp.Tar.UnPackTo Easp.Var("untarfile"), Easp.Var("savepath")
    Easp.Str.JsAlertUrl "解压成功！文件解压到站点的下面位置：" & vbCrLf & Easp.Tar.SavePath, "."
  Case "sitetree" '取站点目录树
    Dim arr1, arr2, obj1, fileList, Root
    Easp.Var("root") = UnEscape(Easp.Var("root"))
    arr1 = Easp.Fso.Dir(Easp.Var("root"))
    Set arr2 = Easp.Json.NewArray
    For i = 0 To Ubound(arr1,2)
      Set obj1 = Easp.Json.NewObject
      obj1.Put "name", Easp.Var("root") & arr1(0,i)
      obj1.Put "type", Easp.IIF(Right(arr1(0,i), 1)="/", "folder", "file")
      arr2.Add obj1
    Next
    fileList = Easp.Encode(arr2)
    Set obj1 = Nothing
    Set arr2 = Nothing
    Easp.PrintEnd fileList
End Select 
%>