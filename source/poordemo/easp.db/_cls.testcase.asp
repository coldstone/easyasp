<!--#include file="../../easyasp/easp.asp" -->
<%
'=========================
'EasyASP测试用例功能类文件
'Author:coldstone
'Update time: 2014-02-05
'=========================

'设置Easp数据源：
' 设置默认的数据源
Easp.Db.SetConnection "default", 1, "_sample.mdb", ""
' 或者
'Easp.Db.SetConn "ACCESS", "_sample.mdb", ""


Dim TestCase
Set TestCase = New EasyASP_TestCase

'测试类
Class EasyASP_TestCase
  Public Db

  Private Sub Class_Initialize()
    Set Db = New EasyAsp_TestCase_Db
  End Sub
  Private Sub Class_Terminate()
    Set Db = Nothing
  End Sub
  
  '打印脚本运行时间
  Public Function ShowScriptTime()
    Easp.print "<small style=""color:#ccc"">页面执行时间 " & Easp.GetScriptTime() & " 秒, 数据库查询 " & Easp.Db.QueryTimes & " 次</small>"
  End Function
End Class

'数据库测试类
Class EasyAsp_TestCase_Db

  '显示记录集
  Sub ShowRecordSet(ByVal rs, ByVal num)
    Dim s_tmp, field, i, n
    n = Easp.IIF(IsNumeric(num),num,0)
    s_tmp = "<style>.easp-test-table{width:100%;margin:0;padding:0;border-collapse:collapse;border:1px solid #555;border-bottom:none;}.easp-test-table th{background-color:#EEE;white-space:nowrap;}.easp-test-table thead th{background-color:#CCC;}.easp-test-table th,.easp-test-table td{font-size:12px;border:1px solid #999;padding:4px;word-break:break-all;}</style>"
    s_tmp = s_tmp & "<table class=""easp-test-table"">"
    If Easp.Has(rs) Then
      s_tmp = s_tmp & "<tr>"
      For Each field In rs.Fields
        s_tmp = s_tmp & "<th>" & field.Name & "</th>"
      Next
      s_tmp = s_tmp & "</tr>"
      If n>rs.RecordCount Then n = rs.RecordCount
      For i = 0 To n-1
        If rs.Eof Then Exit For
        s_tmp = s_tmp & "<tr>"
        For Each field In rs.Fields
          s_tmp = s_tmp & "<td>" & Easp.IIF(TypeName(field.Value)="Byte()", "[ 二进制文件 ]", field.Value) & "</td>"
        Next
        s_tmp = s_tmp & "</tr>"
        rs.MoveNext
      Next
    Else
      s_tmp = s_tmp & "<tr><td>没有数据</td></tr>"
    End If
    s_tmp = s_tmp & "</table>"
    Easp.println s_tmp
  End Sub
  '生成<select>下拉框的option选项
  '参数：@sql - 获取数据的sql语句，第1个字段必须为ID，第2个必须为名称，第3个必须为父级ID
  '     @defaultValue - 默认选中(checked)的项目的ID
  '     @rootParentValue - 根父ID，即只显示此级别及以下的分类
  '返回：<optoin value="...">...</optoin>...不含最外层的<select>标签
  Function GetSelectOptions(ByVal sql, ByVal defaultValue, ByVal rootParentValue)
    Dim s_tmp, rs
    s_tmp = "<option value="""">请选择...</option>"
    Set rs = Easp.Db.Sel(sql)
    '调用递归方法生成各层级的项目
    '(如果最后一个参数为空，则一级分类前不显示左侧前导符，如果需要请设置为和第4个参数一样)
    s_tmp = s_tmp & CreateSelectOption(defaultValue, rs, rootParentValue, "&nbsp;┃", "&nbsp;┃")
    GetSelectOptions = s_tmp
    Easp.db.Close(rs)
  End Function
  '递归生成无限级的树型select中的某级option
  '参数：@defaultValue - 默认选中(checked)的项目的ID
  '     @rs           - 包含全部记录的recordset对象
  '     @parentValue  - 此次要递归处理的父ID
  '     @strSplit     - 默认的在下级名称前要加上的符号（缩进）
  '     @strTab       - 上级名称前已经存在的符号
  '返回：<optoin value="...">...</optoin>
  Private Function CreateSelectOption(ByVal defaultValue, ByVal rs, ByVal parentValue, ByVal strSplit, ByVal strTab)
    Dim s_tmp, rs_tmp, s_split, i_rsCount, i, isLast
    '复制记录集
    Set rs_tmp = rs.Clone()
    '过滤记录集，只取父ID为此次传入的父ID值的记录
    rs_tmp.Filter = rs.Fields(2).Name & "='" & parentValue & "'"
    '取记录集总数
    i_rsCount = rs_tmp.RecordCount
    i = 1 '此变量用于确定当前记录是否最后一条记录
    While Not rs_tmp.Eof
      isLast = (i = i_rsCount)  '判断是否是最后一条记录
      '开始生成标签
      s_tmp = s_tmp & "<option value=""" & rs_tmp(0) & """"
      '如果是默认值添加selected属性
      s_tmp = s_tmp & Easp.IfThen(Easp.Str.IsSame(defaultValue, rs_tmp(0)), " selected")
      '如果strTab为空，则一级分类前不要前导符
      If strTab = "" Then
        s_tmp = s_tmp & ">" & rs_tmp(1) & "</option>"
      Else
        '如果上级分类是最后一条记录（包含隐藏左侧前导符竖线标志），则不显示上级分类左侧的前导符竖线
        s_split = Replace(Left(strTab,Len(strTab)-1), "┊", "&nbsp;&nbsp;&nbsp;")
        '如果是本级最后一条记录，则显示前导符为结束标志
        s_tmp = s_tmp & ">" & s_split & Easp.IIF(isLast,"┖ ","┠ ") & rs_tmp(1) & "</option>"
        '如果是本级最后一条记录，则将隐藏左侧前导符竖线标志传到下级处理
        strTab =  Easp.IIF(isLast, Left(strTab,Len(strTab)-1) & "┊", strTab)
      End If
      '递归处理下一级分类，传入本级的ID、将本级的strTab加上一次strSplit
      s_tmp = s_tmp & CreateSelectOption(defaultValue, rs, rs_tmp(0), strSplit, strTab & strSplit)
      rs_tmp.MoveNext
      i = i + 1
    Wend
    Easp.Db.Close(rs_tmp)
    CreateSelectOption = s_tmp
  End Function
End Class
%>