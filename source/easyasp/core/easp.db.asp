<%
'######################################################################
'## easp.db.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Database Control Class
'## Version     :   3.0
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2014-06-28 1:47:35
'## Description :   Database controler
'##
'######################################################################

Class EasyASP_Db

  Private o_conn, o_connections, o_pager
  Private s_pageParam, s_insSeparator
  Private i_transLevel, i_queryTimes
  Private i_pageIndex, i_pageSize, i_pageCount, i_recordCount, i_rsSize, i_maxRow, i_minRow

  '构造方法
  Private Sub Class_Initialize()
    Easp.Error("error-db-conn") = Easp.Lang("error-db-conn")
    Easp.Error("error-db-noconn") = Easp.Lang("error-db-noconn")
    Easp.Error("error-db-execute") = Easp.Lang("error-db-execute")
    Easp.Error("error-db-batchselect") = Easp.Lang("error-db-batchselect")
    Easp.Error("error-db-executebatch") = Easp.Lang("error-db-executebatch")
    Easp.Error("error-db-paramarray") = Easp.Lang("error-db-paramarray")
    Easp.Error("error-db-select") = Easp.Lang("error-db-select")
    Easp.Error("error-db-getrecordset") = Easp.Lang("error-db-getrecordset")
    Easp.Error("error-db-insert") = Easp.Lang("error-db-insert")
    Easp.Error("error-db-insertbatch") = Easp.Lang("error-db-insertbatch")
    Easp.Error("error-db-delete") = Easp.Lang("error-db-delete")
    Easp.Error("error-db-deletebatch") = Easp.Lang("error-db-deletebatch")
    Easp.Error("error-db-update") = Easp.Lang("error-db-update")
    Easp.Error("error-db-updatebatch") = Easp.Lang("error-db-updatebatch")
    Set o_connections = Server.CreateObject("Scripting.Dictionary")
    o_connections.CompareMode = 1
    '记录默认连接的事务层次
    i_transLevel  = 0
    i_queryTimes  = 0
    s_pageParam   = "page"
    s_insSeparator= ","
    i_pageSize    = 25
    i_pageIndex   = 0
    i_pageCount   = 0
    i_recordCount = 0
    i_rsSize      = 0
    i_maxRow      = 0
    i_minRow      = 0
    Set o_pager   = Server.CreateObject("Scripting.Dictionary")
    o_pager.CompareMode = 1
    SetPager "", "{first}{prev}{liststart}{list}{listend}{next}{last} {jump}", Array("jump:select", "jumplong:0")
    SetPager "bootstrap", "{first}{prev}{list}{next}{last}", Array("listtype:ul", "listclass:pagination pagination-sm", "currentclass:active")
    SetPager "bootstrap.pager", "{prev}{next}", Array("listtype:ul", "currentclass:active", "prev:Previous", "next:Next")
    SetPager "bootstrap.pagerside", "{prev}{next}", Array("listtype:ul", "currentclass:active", "prevclass:previous", "nextclass:next", "prev:&larr; Older", "next:Newer &rarr;")
  End Sub
  '析构方法
  Private Sub Class_Terminate()
    Dim i_level, i
    '释放默认connection对象
    If TypeName(o_conn) = "Connection" Then
      On Error Resume Next
      If i_transLevel>0 Then
        For i = 1 To i_transLevel
          o_conn.RollbackTrans
        Next
      End If
      Err.Clear
      On Error GoTo 0
      If o_conn.State = 1 Then o_conn.Close()
      Set o_conn = Nothing
    End If
    Set o_connections = Nothing
    Set o_pager = Nothing
  End Sub

  '属性：获取数据库的操作次数，只读
  Public Property Get QueryTimes()
    QueryTimes = i_queryTimes
  End Property

  '设置分页标识URL参数
  Public Property Let PageParam(ByVal string)
    s_pageParam = string
  End Property
  '设置和读取分页每页数量
  Public Property Let PageSize(ByVal sizeNumber)
    i_pageSize = sizeNumber
  End Property
  Public Property Get PageSize()
    PageSize = i_pageSize
  End Property
  '读取分页记录集总记录数
  Public Property Get PageRecordCount()
    PageRecordCount = i_recordCount
  End Property
  '读取分页记录集总页数
  Public Property Get PageCount()
    PageCount = i_pageCount
  End Property
  '读取分页记录集当前页码
  Public Property Get PageIndex()
    PageIndex = i_pageIndex
  End Property
  '读取分页记录集当前页记录数
  Public Property Get PageCurrentSize()
    PageCurrentSize = i_rsSize
  End Property
  '读取分页记录集当前页最小记录行号
  Public Property Get PageMinRow()
    PageMinRow = i_minRow
  End Property
  '读取分页记录集当前页最大记录行号
  Public Property Get PageMaxRow()
    PageMaxRow = i_maxRow
  End Property
  '设置Ins和Insert方法字段之间的分隔符，默认为 ,
  Public Property Let InsertSeparator(ByVal string)
    s_insSeparator = string
  End Property

  '生成数据库连接字符串
  '参数：  @dbType    - 数据库类型
  '       @strDB     - 数据库名称
  '       @strServer - 服务器地址及认证信息
  '返回：  ADODB.Connection对象
  Public Function OpenConnection(ByVal dbType, ByVal strDB, ByVal strServer)
    Dim ConnStr, objConn, s, u, p, port
    s = "" : u = "" : p = "" : port = ""
    '分离服务器地址、用户名、密码和端口号
    If Instr(strServer,"@")>0 Then
      s = Trim(Mid(strServer,InstrRev(strServer,"@")+1))
      u = Trim(Left(strServer,InstrRev(strServer,"@")-1))
      If Instr(s,":")>0 Then : port = Trim(Mid(s,Instr(s,":")+1)) : s = Trim(Left(s,Instr(s,":")-1))
      If Instr(u,":")>0 Then : p = Trim(Mid(u,Instr(u,":")+1)) : u = Trim(Left(u,Instr(u,":")-1))
    Else
      If Instr(strServer,":")>0 Then
        u = Trim(Left(strServer,Instr(strServer,":")-1))
        p = Trim(Mid(strServer,Instr(strServer,":")+1))
      Else
        p = Trim(strServer)
      End If
    End If
    Select Case dbType
      Case "MSSQL"
        ConnStr = "Provider=sqloledb;Data Source=" & s & Easp.IfThen(Easp.Has(port), "," & port) & ";Initial Catalog="&strDB&";User Id="&u&";Password="&p&";"
      Case "ACCESS"
        Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
        ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
      Case "MYSQL"
        '服务器需要安装MySQL ODBC驱动，下载地址 http://dev.mysql.com/downloads/connector/odbc/3.51.html
        If port = "" Then port = "3306"
        ConnStr = "Driver={MySQL ODBC 3.51 Driver};Server="&s&";Port="&port&";charset=UTF8;Database="&strDB&";User="&u&";Password="&p&";Option=3;"
    End Select
    'Easp.Console ConnStr
    Set OpenConnection = CreateConnection(ConnStr)
  End Function

  '建立数据库连接对象
  '参数：  @ConnStr  - 数据库连接字符串
  '返回：  ADODB.Connection对象
  Public Function CreateConnection(ByVal ConnStr)
    On Error Resume Next
    Set CreateConnection = Server.CreateObject("ADODB.Connection")
    'Easp.Console connstr
    CreateConnection.Open ConnStr
    '如果打开连接成功
    If Err.number <> 0 Then
      If Request.ServerVariables("LOCAL_ADDR") = Request.ServerVariables("REMOTE_ADDR") Then
        Easp.Error.Detail = ", (""" & ConnStr & """)"
        Easp.Error.FunctionName = "Easp.Db.CreatConnection (easp.db.asp, line 130)"
      End If
      Easp.Error.Raise "error-db-conn"
    End If
  End Function

  '设置Connection对象
  '参数：  @connectionName - 连接名称
  '       @dbType    - 数据库类型
  '       @strDB     - 数据库名称
  '       @strServer - 服务器地址及认证信息
  '返回：  ADODB.Connection对象
  Public Sub SetConnection(ByVal connectionName, ByVal dbType, ByVal strDB, ByVal strServer)
    dbType = UCase(Cstr(dbType))
    If IsNumeric(dbType) Then dbType = Array("MSSQL","ACCESS","MYSQL")(dbType)
    o_connections(connectionName) = Array(dbType, strDB, strServer)
  End Sub
  
  '设置默认Connection对象
  '参数：  @dbType    - 数据库类型
  '       @strDB     - 数据库名称
  '       @strServer - 服务器地址及认证信息
  Public Sub SetConn(ByVal dbType, ByVal strDB, ByVal strServer)
    Call SetConnection("default", dbType, strDB, strServer)
  End Sub
  
  '取得Connection对象
  '参数：  @connectionName - 连接名称
  Public Function GetConnection(ByVal connectionName)
    Dim o, t_start
    t_start = Timer
    '如果连接对象已配置
    If o_connections.Exists(connectionName) Then
      '取得连接字符串要素
      o = o_connections(connectionName)
      Set GetConnection = OpenConnection(o(0),o(1),o(2))
      If Easp.Console.ShowSql Then Easp.Console "连接到数据库(" & GetTypeVersion(GetConnection) & ")：" & Mid(o(1), InStrRev(o(1), "/")+1) & Easp.IfThen(o(0)<>"ACCESS" ,"@" & Trim(Mid(o(2),InstrRev(o(2),"@")+1))) & "， 执行时间：" & Easp.GetScriptTimeByTimer(t_start) & "s"
    Else
      '没有找到此名称的连接对象
      Easp.Error.FunctionName = "Easp.Db.GetConnection"
      Easp.Error.Detail = connectionName
      Easp.Error.Raise "error-db-noconn"
    End If
  End Function

  '取得默认Connection对象
  Public Function GetConn()
    OpenConn()
    Set GetConn = o_conn
  End Function

  '接管外部Connection对象为默认连接
  Public Property Let DefaultConn(ByRef conn)
    If TypeName(conn) = "Connection" Then
      If conn.State = 1 Then
        Set o_conn = conn
        If Easp.Console.ShowSql Then Easp.Console "接管外部连接为默认数据库(" & GetTypeVersion(conn) & ")"
      End If
    End If
  End Property
  
  '打开默认Connection连接
  Private Sub OpenConn()
    Dim b_opened
    If TypeName(o_conn) = "Connection" Then
      If o_conn.State = 1 Then b_opened = True
    End If
    If Not b_opened Then
      Set o_conn = GetConnection("default")
    End If
  End Sub
  
  '取得数据库类型
  Public Function GetType(ByRef conn)
    Dim dbms, s_type
    dbms = conn.Properties("DBMS Name")
    If Easp.Str.IsSame(dbms, "MS Jet") Then
      s_type = "ACCESS"
    ElseIf Easp.Str.IsSame(dbms, "Microsoft SQL Server") Then
      s_type = "MSSQL"
    ElseIf Easp.Str.IsSame(dbms, "MySQL") Then
      s_type = "MYSQL"
    Else
      s_type = dbms
    End If
    GetType = s_type
  End Function
  '取得默认连接数据库类型
  Public Function [Type]()
    OpenConn()
    [Type] = GetType(o_conn)
  End Function
  
  '取得数据库的版本号
  Public Function GetVersion(ByRef conn)
    GetVersion = conn.Properties("DBMS Version")
  End Function
  '取得默认连接数据库的版本号
  Public Function Version()
    OpenConn()
    Version = o_conn.Properties("DBMS Version")
  End Function
  '取得更清晰的版本号
  Public Function GetTypeVersion(ByRef conn)
    Dim s_type, s_ver
    s_ver = GetVersion(conn)
    Select Case GetType(conn)
      Case "ACCESS" s_type = "Microsoft Access "
        Select Case Left(s_ver, 2)
          Case "04" s_ver = "2000-2003"
          Case "12" s_ver = "2007"
        End Select
      Case "MSSQL"  s_type = "Microsoft SQL Server "
        Select Case Left(s_ver, 8)
          Case "12.00.20" s_ver = "2014 RTM"
          Case "11.00.30" s_ver = "2012 Service Pack 1"
          Case "11.00.21" s_ver = "2012 RTM"
          Case "10.50.40" s_ver = "2008 R2 Service Pack 2"
          Case "10.50.25" s_ver = "2008 R2 Service Pack 1"
          Case "10.50.16" s_ver = "2008 R2 RTM"
          Case "10.00.55" s_ver = "2008 Service Pack 3"
          Case "10.00.40" s_ver = "2008 Service Pack 2"
          Case "10.00.25" s_ver = "2008 Service Pack 1"
          Case "10.00.16" s_ver = "2008 RTM"
          Case "9.00.500" s_ver = "2005 Service Pack 4"
          Case "9.00.403" s_ver = "2005 Service Pack 3"
          Case "9.00.304" s_ver = "2005 Service Pack 2"
          Case "9.00.204" s_ver = "2005 Service Pack 1"
          Case "9.00.139" s_ver = "2005 RTM"
          Case "8.00.203" s_ver = "2000 Service Pack 4"
          Case "8.00.760" s_ver = "2000 Service Pack 3"
          Case "8.00.534" s_ver = "2000 Service Pack 2"
          Case "8.00.384" s_ver = "2000 Service Pack 1"
          Case "8.00.194" s_ver = "2000 RTM"
        End Select
      Case "MYSQL"  s_type = "MySQL Server "
      Case Else
        s_type = conn.Properties("DBMS Name") & " "
    End Select
    GetTypeVersion = s_type & s_ver
  End Function
  '取得更清晰的版本号
  Public Function TypeVersion()
    OpenConn()
    TypeVersion = GetTypeVersion(o_conn)
  End Function

  '执行SQL原型
  Private Function ExecuteSql(ByRef conn, ByVal sql, ByVal executeType)
    Dim cmd, match, matchCount, i, affectedRows, currentCursor, sTimer
    Dim o_sql, queryType, sqlParam, sqlParamType, sqlParamSize, outSql
    Dim param(), paramValue(), paramType(), paramSize(), inOrOut()
    sTimer = Timer()
    '判断是否是查询语句
    If Easp.Str.IsSame(Left(sql, 7),"select ") Then queryType = 1
    '判断是否调用存储过程或函数
    If Easp.Str.IsInList("call ,exec ", Left(sql, 5)) Then queryType = 4
    sql = ReplaceStasicParameter(sql)
    sql = ReplaceNewId(sql)
    outSql = sql
    o_sql = sql
    '查找参数标签
    Set match = Easp.Str.Match(sql, "\{(.+?)\}")
    matchCount = match.Count-1
    '如果sql中包含参数
    If matchCount>=0 Then
      '定义参数名，参数值，数据类型，参数类型数组
      ReDim param(matchCount)
      ReDim paramValue(matchCount)
      ReDim paramType(matchCount)
      ReDim paramSize(matchCount)
      ReDim inOrOut(matchCount)
      For i = 0 To matchCount
        '取出参数标签内容
        sqlParam = match(i).SubMatches(0)
        '取参数名
        param(i) = Easp.Str.GetColonName(sqlParam)
        'Easp.Println "p & i::" & param(i) & "---" & inOrOut(i)
        If Left(param(i),2)="@@" Then
        '如果既是输入参数又是输出参数
          param(i) = Mid(param(i), 3)
          inOrOut(i) = 3
        ElseIf Left(param(i),1)="@" Then
        '如果是输出参数
          param(i) = Mid(param(i), 2)
          inOrOut(i) = 2
        Else
          inOrOut(i) = 1
        End If
        'Easp.Println "p & i::" & param(i) & "---" & inOrOut(i)
        sqlParamType = Easp.Str.GetColonValue(sqlParam)
        paramSize(i) = 8000
        '取数据类型，不设置默认为varchar类型
        If Instr(sqlParamType, "(") Then
          paramType(i) = GetParameterType(Easp.Str.GetName(sqlParamType, "("))
          sqlParamSize = Easp.Str.GetName(Easp.Str.GetValue(sqlParamType, "("), ")")
          If IsNumeric(sqlParamSize) Then paramSize(i) = sqlParamSize
        Else
          paramType(i) = GetParameterType(sqlParamType)
        End If
        '取参数的原始值，并处理静态标签嵌套
        paramValue(i) = Easp.Var(param(i))
        '替换输出SQL语句中的参数为值
        outSql = Replace(outSql, match(i), FormatValue(paramValue(i), paramType(i)))
      Next
    End If
    Set match = Nothing
    '仅输出SQL语句
    If executeType = -1 Then
      ExecuteSql = outSql
      Exit Function
    End If
    '在控制台中输出SQL语句
    If Easp.Console.ShowSql Then Easp.Console outSql
    sql = Easp.Str.Replace(sql, "\{(.+?)\}", "?")
    '定义Command对象
    Set cmd = Server.CreateObject("ADODB.Command")
    With cmd
      .ActiveConnection = conn
      .CommandText = sql
      .CommandType = 1
      .Prepared = True
      If queryType = 4 Then
      '如果是存储过程
        Dim spName, sp, spParams, spParam, spParamValue
        Dim spParamType, spParamSize, spParamInorOut, j
        spName = Split(o_sql)(1)
        .CommandText = spName
        .CommandType = 4
        '添加返回值参数
        .Parameters.append .CreateParameter("sp_return", 3, 4)
        '解析参数
        spParam = Trim(Easp.Str.GetValue(o_sql, spName))
        If Easp.Has(spParam) Then
        '如果存储过程有参数
          spParams = Split(spParam, ",")
          'Easp.PrintlnString spParams
          For i = 0 To UBound(spParams)
            '处理非{var}参数值
            spParam = Trim(spParams(i))
            If Left(spParam, 1) = "'" And Right(spParam, 1) = "'" Then
              spParamValue = Mid(spParam, 2, Len(spParam) -2)
            End If
            spParamType = 200
            spParamInorOut = 1
            '处理{var}参数
            If Left(spParam, 1) = "{" And Right(spParam, 1) = "}" Then
              '取出参数名
              spParam = Mid(spParam, 2, Len(spParam) -2)
              spParam = Easp.Str.GetColonName(spParam)
              'Easp.Println spParam
              '取出参数的值和类型
              If matchCount>=0 Then
                For j = 0 To matchCount
                  If Easp.Str.IsInList(param(j) & ",@" & param(j) & ",@@" & param(j), spParam) Then
                  '匹配到参数
                    spParamValue = paramValue(j)
                    spParamType = paramType(j)
                    spParamSize = paramSize(j)
                    spParamInorOut = inOrOut(j)
                    Exit For
                  End If
                Next
              End If
            End If
            If spParamInorOut = 1 Or spParamInorOut = 3 Then
            '输入参数
              '如果是数值型/日期型而且为空，则输入NULL值
              If InStr(",20,11,6,5,7,135,3,131,4,2,16,128,205,204,", "," & spParamType & ",") > 0 And Easp.IsN(spParamValue) Then spParamValue = Null
              'If IsNumeric(spParamValue) And spParamType = 200 Then spParamType = 5
              'If IsDate(spParamValue) And spParamType = 200 Then spParamType = 135
              .Parameters.Append .CreateParameter(spParam, spParamType, spParamInorOut, spParamSize, spParamValue)
              'Easp.PrintlnString Array(spParam, spParamType, spParamInorOut, spParamSize, spParamValue)
            Else
            '输出参数
              .Parameters.Append .CreateParameter(spParam, spParamType, 2, spParamSize)
              'Easp.PrintlnString Array(spParam, spParamType, 2, spParamSize)
            End If
          Next
        End If
        Dim rsState, list, pKey, outParams, returnValue
        conn.CursorLocation = 3
        Set sp = .Execute(affectedRows, , 4)
        rsState = sp.State
        'Easp.Println "sp.State::" & sp.State
        'Easp.Println "sp.RecordCount::" & sp.RecordCount
        '关闭了记录集才能取输出参数
        If rsState = 1 Then sp.Close
        Set List = Easp.Json.NewObject
        '取输出参数
        Set outParams = Easp.Json.NewObject
        For i = 0 To .Parameters.Count -1
          If i = 0 Then returnValue = .Parameters(i).Value
          'If Left(.Parameters(i).Name, 1) = "@" Then
          outParams(.Parameters(i).Name) = .Parameters(i).Value
        Next
        '写入受影响的行数
        List.Put "rows", affectedRows
        '写入返回值
        List.Put "return", returnValue
        '写入输出参数
        List.Put "out", outParams.GetDictionary
        Set outParams = Nothing
        '写入记录集
        If rsState = 1 Then
          sp.Open
          List.Put "rs", sp.Clone
        Else
          List.Put "rs", Null
        End If
        Set ExecuteSql = List.GetDictionary
        Set List = Nothing
        Easp.Db.Close(sp)
        .ActiveConnection = Nothing
        i_queryTimes = i_queryTimes + 1
        If Easp.Console.ShowSql And Easp.Console.ShowSqlTime Then
          Easp.Console "(" & Easp.Lang("db-query-spend") & "：" & Easp.GetScriptTimeByTimer(sTimer) & "s， " & Easp.Lang("db-return") & "：" & returnValue & ")"
        End If
      Else
        If matchCount>=0 Then
        '如果sql中包含参数
          For i = 0 To matchCount
            '如果是数值型/日期型而且为空，则输入NULL值
            If InStr(",20,11,6,5,7,135,3,131,4,2,16,128,205,204,", ","&paramType(i)&",") > 0 And Easp.IsN(paramValue(i)) Then paramValue(i) = Null
            'If IsNumeric(paramValue(i)) And paramType(i) = 200 Then paramType(i) = 5
            'If IsDate(paramValue(i)) And paramType(i) = 200 Then paramType(i) = 135
            .Parameters.Append .CreateParameter(param(i), paramType(i), 1, paramSize(i), paramValue(i))
          Next
        End If
        If queryType = 1 Or executeType = 1 Then
        '如果要返回记录集
          'currentCursor = conn.CursorLocation
          conn.CursorLocation = 3 '游标服务位置为client时才能返回正确的RecordCount
          Set ExecuteSql = .Execute
          .ActiveConnection = Nothing
          i_queryTimes = i_queryTimes + 1
          If Easp.Console.ShowSql And Easp.Console.ShowSqlTime Then
            Dim i_rcount
            i_rcount = 0
            If ExecuteSql.State = 1 Then i_rcount = ExecuteSql.RecordCount
            Easp.Console "(" & Easp.Lang("db-query-spend") & "：" & Easp.GetScriptTimeByTimer(sTimer) & "s， " & Easp.Lang("db-record-count") & "：" & i_rcount & ")"
          End If
          'conn.CursorLocation = currentCursor
        ElseIf executeType = 2 Then
        '如果仅返回执行成功与否
          .Execute affectedRows, , 129
          .ActiveConnection = Nothing
          ExecuteSql = (affectedRows>0)
          i_queryTimes = i_queryTimes + 1
          If Easp.Console.ShowSql And Easp.Console.ShowSqlTime Then
            Easp.Console "(" & Easp.Lang("db-query-spend") & "：" & Easp.GetScriptTimeByTimer(sTimer) & "s， " & Easp.IIF(ExecuteSql, Easp.Lang("db-query-success"), Easp.Lang("db-query-fail")) & "， " & Easp.Lang("db-affected-rows") & "：" & affectedRows & ")"
          End If
        Else'If executeType = 0 Then
        '返回受影响的行数
          .Execute affectedRows, , 129
          .ActiveConnection = Nothing
          ExecuteSql = affectedRows
          i_queryTimes = i_queryTimes + 1
          If Easp.Console.ShowSql And Easp.Console.ShowSqlTime Then
            Easp.Console "(" & Easp.Lang("db-query-spend") & "：" & Easp.GetScriptTimeByTimer(sTimer) & "s， " & Easp.Lang("db-affected-rows") & "：" & affectedRows & ")"
          End If
        End If
      End If
    End With
    Set cmd = Nothing
  End Function

  '仅显示出sql
  Public Function ToSql(ByVal sql)
    ToSql = ExecuteSql("", sql, -1)
  End Function
  '仅显示出批量sql
  Public Function ToSqlBatch(ByVal sql)
    Dim a_sql, i, s_tmp
    a_sql = GetBatchArray(sql)(0)
    For i = 0 To UBound(a_sql)
      If i>0 Then s_tmp = s_tmp & vbCrLf
      s_tmp = s_tmp & ExecuteSql("", a_sql(i), -1) & ";"
    Next
    ToSqlBatch = s_tmp
  End Function

  '抛出错误调试信息
  Private Sub CheckError(ByVal name, ByRef e, ByRef conn, ByRef funName, ByVal detail)
    Dim b_tmp
    If IsObject(e) Then
      If e.Number <> 0 Or conn.Errors.Count > 0 Then b_tmp = True
    Else
      If e = "" Then b_tmp = True
    End If
    If b_tmp Then
      'If Easp.Debug Then
        Easp.Error.SetErrors e, conn, Null
        Easp.Error.FunctionName = "Easp.Db." & funName
        Easp.Error.Detail = detail & Easp.Lang("db-see-sql-in-console")
        Easp.Error.Raise "error-db-" & name
      'End If
    End If
  End Sub
  
  '执行SQL语句,返回记录集(R)或受影响的行数(CUD)
  Public Function Execute(ByRef conn, ByVal sql)
    On Error Resume Next
    '判断是否是查询语句
    If Easp.Str.IsSame(Left(sql, 7),"select ") Or Easp.Str.IsInList("call ,exec ", Left(sql, 5)) Then
      Set Execute = ExecuteSql(conn, sql, 1)
    Else
      Execute = ExecuteSql(conn, sql, 0)
    End If
    CheckError "execute", Err, conn, "Execute", sql
  End Function
  
  '用默认Connection执行SQL语句,返回记录集(R)或受影响的行数(CUD)
  Public Function Exec(ByVal sql)
    On Error Resume Next
    OpenConn()
    If Easp.Str.IsSame(Left(sql, 7),"select ") Or Easp.Str.IsInList("call ,exec ", Left(sql, 5)) Then
      Set Exec = ExecuteSql(o_conn, sql, 1)
    Else
      Exec = ExecuteSql(o_conn, sql, 0)
    End If
    CheckError "execute", Err, o_conn, "Exec", sql
  End Function
  '用默认Connection执行SQL语句，返回记录集(R)或成功与否(CUD)
  Public Function Query(ByVal sql)
    On Error Resume Next
    OpenConn()
    If Easp.Str.IsSame(Left(sql, 7),"select ") Or Easp.Str.IsInList("call ,exec ", Left(sql, 5)) Then
      Set Query = ExecuteSql(o_conn, sql, 1)
    Else
      Query = ExecuteSql(o_conn, sql, 2)
    End If
    CheckError "execute", Err, o_conn, "Query", sql
  End Function

  '批量执行SQL语句
  Public Function ExecuteBatch(ByRef conn, ByVal sql)
    On Error Resume Next '## Do not delete or comment
    Err.Clear
    Dim a_sql, i, i_result
    i_result = 0
    If Easp.Str.IsSame(Left(sql, 7),"select ") Or Easp.Str.IsInList("call ,exec ", Left(sql, 5)) Then
      CheckError "batchselect", "", conn, "ExecuteBatch", sql
      Exit Function
    End If
    a_sql = GetBatchArray(sql)(0)
    '开始事务
    BeginTrans(conn)
    For i = 0 To UBound(a_sql)
      i_result = i_result + ExecuteSql(conn, a_sql(i), 0)
    Next
    If Err.number = 0 And conn.Errors.Count = 0 Then
      '提交事务
      CommitTrans(conn)
    Else
      RollbackTrans(conn)
      CheckError "executebatch", Err, conn, "ExecuteBatch / ExecBatch", sql
    End If
    ExecuteBatch = i_result
  End Function
  '用默认Connection批量执行SQL语句
  Public Function ExecBatch(ByVal sql)
    OpenConn()
    ExecBatch = ExecuteBatch(o_conn, sql)
  End Function

  '根据数组参数返回多个sql语句数组
  Private Function GetBatchArray(ByVal sql)
    Dim match, matchCount, a_tmp, a_tmplen
    Dim sqlParam, param(), paramType(), paramBatch()
    Dim i, j, k, sqlCount, s_sql, a_sql()
    '替换静态标签
    sql = ReplaceStasicParameter(sql)
    '查找参数标签
    Set match = Easp.Str.Match(sql, "\{(.+?)\}")
    matchCount = match.Count-1
    sqlCount = 0
    '如果sql中包含参数
    If matchCount>=0 Then
      '定义参数名，数据类型数组
      ReDim param(matchCount)
      ReDim paramType(matchCount)
      j = 0
      For i = 0 To matchCount
        '取出参数标签内容
        sqlParam = match(i).SubMatches(0)
        '取参数名
        param(i) = Easp.Str.GetColonName(sqlParam)
        '取数据类型
        paramType(i) = GetParameterType(Easp.Str.GetColonValue(sqlParam))
        '如果参数是数组
        a_tmp = Easp.Var(param(i)&"_array")
        If IsArray(a_tmp) Then
          If Ubound(a_tmp) > 0 Then
            '取第一个数组参数的数量
            a_tmplen = UBound(a_tmp)
            If sqlCount = 0 Then sqlCount = a_tmplen
            '判断所有为数组的参数所包含的数组数量是否一致
            If sqlCount > 0 And sqlCount <> a_tmplen Then
              Easp.Error.FunctionName = "Easp.Db.GetBatchArray"
              Easp.Error.Detail = param(i)
              Easp.Error.Raise "error-db-paramarray"
              Exit Function
            End If
            '将所有是数组的参数用一个数组保存起来
            ReDim Preserve paramBatch(2,j)
            paramBatch(0,j) = match(i)
            paramBatch(1,j) = param(i)
            paramBatch(2,j) = paramType(i)
            j = j + 1
          End If
        End If
      Next
      ReDim a_sql(sqlCount)
      '执行多条SQL语句
      If sqlCount > 0 Then
        For i = 0 To sqlCount
          '取原始SQL
          s_sql = sql
          For j = 0 To UBound(paramBatch,2)
            '循环替换数组参数为数组内的值
            s_sql = Replace(s_sql, paramBatch(0,j), "{" & paramBatch(1,j) & "_array_" & i & ":" & paramBatch(2,j) & "}")
          Next
          a_sql(i) = ReplaceNewId(s_sql)
        Next
      Else
        a_sql(0) = ReplaceNewId(sql)
      End If
    End If
    GetBatchArray = Array(a_sql,match.Count)
    Set match = Nothing
  End Function

  '取得记录集
  Public Function [Select](ByRef conn, ByVal sql)
    On Error Resume Next
    Set [Select] = ExecuteSql(conn, sql, 1)
    CheckError "select", Err, conn, "Select", sql
  End Function
  '用默认Connection取得记录集
  Public Function Sel(ByVal sql)
    On Error Resume Next
    OpenConn()
    Set Sel = ExecuteSql(o_conn, sql, 1)
    CheckError "select", Err, o_conn, "Sel", sql
  End Function

  Public Function NextRS(ByRef rs)
    Do While True
      Set NextRS = rs.NextRecordset
      If NextRS Is Nothing Then
        Exit Do
      ElseIf TypeName(NextRS) = "Recordset" Then
        If NextRS.State = 1 Then Exit Do
      End If
    Loop
  End Function  

  '取得分页后记录集
  '说明：可以是单表、多表连接或者包含子查询的复杂SQL查询语句，但如果是Access数据
  '     库或者SQL Server 2000及以下版本数据库，则必须满足以下三个条件：
  '      1、sql中Select的第一个字段必须是主键
  '      2、所有参与排序的字段必须在Select出的字段中包含
  '      3、Order By语句中不能出现括号
  '提示：为提高分页效率，MSSQL或MySQL数据库请提前建好索引
  Public Function GetRecordSet(ByRef conn, ByVal sql)
    On Error Resume Next
    Dim rsTmp, s_keySql, b_insideSql, s_pkey, s_order, s_reOrder
    Dim s_tmp, s_sqlNoOrder, s_sqlCount, i_tmp
    Dim s_dbType, s_dbVer, s_sql, s_osql
    Dim t_start : t_start = Timer : s_osql = sql
    b_insideSql = Easp.Console.ShowSql
    Easp.Console.ShowSql = False
    If b_insideSql Then
      Easp.Console "(分页) " & ToSql(sql)
    End If
    '取主键名
    s_keySql = Replace(sql, "Select", "SELECT TOP 1", 1, 1, 1)
    s_keySql = "SELECT * FROM (" & s_keySql & ") AS EasyASP_Pager_Key_Table WHERE 1=0"
    Set rsTmp = ExecuteSql(conn, s_keySql, 1)
    i_queryTimes = i_queryTimes - 1
    s_pkey = rsTmp.Fields(0).Name
    Close(rsTmp)
    'Easp.Console s_pkey
    '取排序字段
    s_tmp = Mid(sql,InStrRev(sql, ")")+1)
    If Easp.Str.IsIn(s_tmp, "order by") Then
      s_order = Mid(s_tmp, InStrRev(s_tmp, "Order by", -1, 1))
      s_sqlNoOrder = Trim(Left(sql, Len(sql)-Len(s_order)))
    Else
      s_order = "ORDER BY " & s_pkey & " ASC"
      s_sqlNoOrder = sql
      sql = sql & " " & s_order
    End If
    '取总记录数
    s_sqlCount = "SELECT COUNT(*) FROM (" & s_sqlNoOrder & ") AS EasyASP_Pager_Count_Table"
    Set rsTmp = ExecuteSql(conn, s_sqlCount, 1)
    i_queryTimes = i_queryTimes - 1
    i_recordCount = rsTmp(0)
    Close(rsTmp)
    If i_recordCount > 0 Then
      i_tmp = i_recordCount/i_pageSize
      i_pageCount = Int(i_tmp) + Easp.IIF(Int(i_tmp)=i_tmp, 0, 1)
      '取当前页码
      i_pageIndex = GetPageIndex()
      If i_pageIndex > i_pageCount Then i_pageIndex = i_pageCount
      '计算记录行号
      i_minRow = i_pageSize * (i_pageIndex-1) + 1
      i_maxRow = i_pageSize * i_pageIndex
      '本页记录数
      i_rsSize = i_pageSize
      If i_maxRow > i_recordCount Then
      '最后一页的最大行号不超过记录总数
        i_maxRow = i_recordCount
        '最后一页的记录数
        i_rsSize = i_maxRow - i_minRow + 1
      End If
      '取得数据库类型及版本
      s_dbType = GetType(conn)
      s_dbVer = GetVersion(conn)
      '按不同类型的数据库来处理分页
      If s_dbType = "MYSQL" Then
      '如果是MySQL数据库，采用limit取分页记录
        s_sql = sql & " limit " & i_minRow - 1 & ", " & i_pageSize
      Else
        If s_dbType = "MSSQL" And Easp.Str.GetName(s_dbVer,".") >= 9 Then
        '如果是SQL Server 2005及以上版本数据库，利用ROW_NUMBER函数取分页记录
          s_sql = "SELECT * FROM (SELECT *,ROW_NUMBER() OVER (" & s_order & ") AS EasyASP_Pager_RowRank FROM (" & s_sqlNoOrder & ") AS EasyASP_Pager_Max_Table) AS EasyASP_Pager_Result_Table WHERE EasyASP_Pager_Result_Table.EasyASP_Pager_RowRank BETWEEN " & i_minRow & " AND " & i_maxRow
        Else
        '如果是Access或者SQL Server 2000及以下版本
          sql = Replace(sql, "Select", "SELECT TOP " & i_maxRow, 1, 1, 1)
          s_sql = "SELECT * From (" & sql & ") AS EasyASP_Pager_Result_Table WHERE " & s_pkey & " IN (SELECT TOP " & i_rsSize & " " & s_pkey & " FROM (" & sql & ") AS EasyASP_Pager_Max_Table " & ReverseOrderBy(s_order) & ") " & s_order & ""
        End If
      End If
      Set GetRecordSet = ExecuteSql(conn, s_sql, 1)
    Else
      i_rsSize = 0
      i_minRow = 0
      i_maxRow = 0
      i_pageIndex = 0
      i_pageCount = 0
      Set GetRecordSet = ExecuteSql(conn, s_osql, 1)
    End If
    '输出执行时间及执行结果
    If b_insideSql And Easp.Console.ShowSqlTime Then
      s_tmp = "(" & Easp.Lang("db-query-spend") & "：" & Easp.GetScriptTimeByTimer(t_start) & "s"
      s_tmp = s_tmp & Easp.Str.Format(Easp.Lang("db-pager-in-console"), Array(i_rsSize, i_minRow, i_maxRow, i_recordCount, i_pageSize, i_pageIndex, i_pageCount))
      Easp.Console s_tmp
    End If
    Easp.Console.ShowSql = b_insideSql
    CheckError "getrecordset", Err, conn, "GetRecordSet", s_osql
  End Function
  '用默认Connection取得分页后记录集
  Public Function GetRS(ByVal sql)
    On Error Resume Next
    OpenConn()
    Set GetRS = GetRecordSet(o_conn, sql)
    CheckError "getrecordset", Err, o_conn, "GetRS", sql
  End Function
  '反转Order By语句
  Private Function ReverseOrderBy(ByVal string)
    Dim s_fields, a_fields, i, s_tmp, s_field, s_sort
    s_fields = Trim(Mid(string,10))
    a_fields = Split(s_fields, ",")
    For i = 0 To UBound(a_fields)
      s_tmp = Trim(a_fields(i))
      If Easp.Str.IsSame(Right(s_tmp, 5), " desc") Then
        s_field = Left(s_tmp, Len(s_tmp) - 5)
        s_sort = " ASC"
      ElseIf Easp.Str.IsSame(Right(s_tmp, 4), " asc") Then
        s_field = Left(s_tmp, Len(s_tmp) - 4)
        s_sort = " DESC"
      Else
        s_field = s_tmp
        s_sort = " DESC"
      End If
      a_fields(i) = s_field & s_sort
    Next
    ReverseOrderBy = "ORDER BY " & Join(a_fields, ", ")
  End Function
  
  '取得当前页码
  '返回：  Int
  Private Function GetPageIndex()
    Dim i_page
    i_page = Easp.Var(s_pageParam)
    i_page = Easp.IIF(isNumeric(i_page) And Easp.Has(i_page), i_page, 1)
    GetPageIndex = Int(i_page)
  End Function

  '生成分页导航链接
  '参数：@html    - 分页导航模板
  '     @config  - 分页导航配置
  '返回：  String
  Public Function Pager(ByVal html, ByRef config)
    'On Error Resume Next
    Dim s_list, i_listStart, i_listEnd, s_first, s_prev, s_next, s_last
    Dim s_jump, i_jumpLong, i_jumpStart, i_jumpEnd, s_jumpValue
    Dim i, j, s_tmp, i_start, i_end, o_cfg, a_pageClass(1), s_moreTmp
    s_tmp = Easp.IfHas(html, o_pager("default_html"))
    Set o_cfg = Server.CreateObject("Scripting.Dictionary")
    o_cfg.CompareMode = 1
    '分页配置参数
    o_cfg("recordcount")  = i_recordCount '总记录数
    o_cfg("pageindex")    = i_pageIndex '当前页码
    o_cfg("pagecount")    = i_pageCount '总页数
    o_cfg("pagesize")     = i_pageSize '每页条数
    o_cfg("rssize")       = i_rsSize '当前页记录数
    o_cfg("minrow")       = i_minRow '当前页最小记录号
    o_cfg("maxrow")       = i_maxRow '当前页最大记录号
    o_cfg("list")         = "*" '页码显示，*代表页码，例如可配置为 第*页
    o_cfg("listtype")     = "div" '分页显示容器，可以是 div 或者 ul，默认为div
    o_cfg("listclass")    = "pager" '分页容器（div或ul）的class样式
    o_cfg("listlong")     = 7 '在分页链接中显示的页码数量
    o_cfg("listsidelong") = 2 '在分页链接两头显示的页码数量，为0则不显示
    o_cfg("pageclass")    = "" '每个页码的class样式
    o_cfg("currentclass") = "current" '当前页的class样式
    o_cfg("disabledclass")= "disabled" '不可用的链接的class样式
    o_cfg("link")         = Easp.ReplaceUrl(s_pageParam, "*") '页码链接地址，其中*代表页码
    o_cfg("first")        = "&laquo;" '首页链接文字
    o_cfg("firstclass")   = "" '首页链接class样式
    o_cfg("prev")         = "&#8249;" '上一页链接文字
    o_cfg("prevclass")    = "" '上一页链接class样式
    o_cfg("next")         = "&#8250;" '下一页链接文字
    o_cfg("nextclass")    = "" '下一页链接class样式
    o_cfg("last")         = "&raquo;" '末页链接文字
    o_cfg("lastclass")    = "" '末页链接class样式
    o_cfg("more")         = "..." '被省略的页码显示为，默认是"..."
    o_cfg("jump")         = "input" '跳转框样式，默认为"input"文本框，可设置为"select"下拉菜单
    o_cfg("jumpplus")     = "" '设置input或select跳转框的标签内属性
    o_cfg("jumpaction")   = "" '跳转时要执行的javascript代码，用*代表页码，默认为跳转到页码链接地址
    o_cfg("jumplong")     = 50 '跳转框为select时下拉菜单包含的页码最大数量，0为全部显示
    '读入配置信息
    '如果为空则读入默认配置
    config = Easp.IfHas(config, o_pager("default_config"))
    '如果配置不是数组则转为数组
    If Easp.Has(config) Then config = Easp.IIF(IsArray(config), config, Array(config,"userconfig:1"))
    If isArray(config) Then
      Dim s_name, s_value
      For i = 0 To Ubound(config)
        s_name = Easp.Str.GetColonName(config(i))
        s_value = Easp.Str.GetColonValue(config(i))
        If Easp.Str.IsInList("recordcount,pageindex,pagecount,pagesize,listlong,listsidelong,jumplong", s_name) Then
          o_cfg(s_name) = Int(s_value)
        Else
          o_cfg(s_name) = s_value
        End If
      Next
    End If
    '计算要显示的页码列表
    i_start = o_cfg("pageindex") - ((o_cfg("listlong") \ 2) + (o_cfg("listlong") Mod 2)) + 1
    i_end = o_cfg("pageindex") + (o_cfg("listlong") \ 2)
    If i_start < 1 Then
      i_start = 1
      i_end = o_cfg("listlong")
    End If
    If i_end > o_cfg("pagecount") Then
      i_start = o_cfg("pagecount") - o_cfg("listlong") + 1
      i_end = o_cfg("pagecount")
      If i_start < 1 Then i_start = 1
    End If
    '生成页码链接列表
    For i = i_start To i_end
      If i = o_cfg("pageindex") Then
        If o_cfg("listtype") = "ul" Then
          s_list = s_list & " <li class=""" & o_cfg("currentclass") & AddHtmlClass(1, o_cfg("pageclass")) & """><a href=""javascript:void(0)"">" & Replace(o_cfg("list"),"*",i) & "</a></li> "
        Else
          s_list = s_list & " <span class=""" & o_cfg("currentclass") & AddHtmlClass(1, o_cfg("pageclass")) & """>" & Replace(o_cfg("list"),"*",i) & "</span> "
        End If
      Else
        If o_cfg("listtype") = "ul" Then
          s_list = s_list & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><a href=""" & Replace(o_cfg("link"),"*",i) & """>" & Replace(o_cfg("list"),"*",i) & "</a></li> "
        Else
          s_list = s_list & " <a href=""" & Replace(o_cfg("link"),"*",i) & """" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & Replace(o_cfg("list"),"*",i) & "</a> "
        End If
      End If
    Next
    '计算分页链接两头显示的页码数量
    If o_cfg("listsidelong")>0 Then
      '生成分页头部链接
      If o_cfg("listsidelong") < i_start Then
        For i = 1 To o_cfg("listsidelong")
          If o_cfg("listtype") = "ul" Then
            i_listStart = i_listStart & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><a href=""" & Replace(o_cfg("link"),"*",i) & """>" & Replace(o_cfg("list"),"*",i) & "</a></li> "
          Else
            i_listStart = i_listStart & " <a href=""" & Replace(o_cfg("link"),"*",i) & """" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & Replace(o_cfg("list"),"*",i) & "</a> "
          End If
        Next
        s_moreTmp = Easp.IfThen(o_cfg("listsidelong") + 1 <> i_start, o_cfg("more"))
        If o_cfg("listtype") = "ul" Then
          i_listStart = i_listStart & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><span>" & s_moreTmp & "</span></li> "
        Else
          i_listStart = i_listStart & " <span" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & s_moreTmp & "</span> "
        End If
      ElseIf o_cfg("listsidelong") >= i_start And i_start > 1 Then
        For i = 1 To (i_start - 1)
          If o_cfg("listtype") = "ul" Then
            i_listStart = i_listStart & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><a href=""" & Replace(o_cfg("link"),"*",i) & """>" & Replace(o_cfg("list"),"*",i) & "</a></li> "
          Else
            i_listStart = i_listStart & " <a href=""" & Replace(o_cfg("link"),"*",i) & """" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & Replace(o_cfg("list"),"*",i) & "</a> "
          End If
        Next
      End If
      '生成分页尾部链接
      If (o_cfg("pagecount") - o_cfg("listsidelong")) > i_end Then
        If o_cfg("listtype") = "ul" Then
          i_listEnd = " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><span>" & o_cfg("more") & "</span></li> "
        Else
          i_listEnd = " <span" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & o_cfg("more") & "</span> "
        End If
        For i = ((o_cfg("pagecount") - o_cfg("listsidelong"))+1) To o_cfg("pagecount")
          If o_cfg("listtype") = "ul" Then
            i_listEnd = i_listEnd & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><a href=""" & Replace(o_cfg("link"),"*",i) & """>" & Replace(o_cfg("list"),"*",i) & "</a></li> "
          Else
            i_listEnd = i_listEnd & " <a href=""" & Replace(o_cfg("link"),"*",i) & """" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & Replace(o_cfg("list"),"*",i) & "</a> "
          End If
        Next
      ElseIf (o_cfg("pagecount") - o_cfg("listsidelong")) <= i_end And i_end < o_cfg("pagecount") Then
        For i = (i_end+1) To o_cfg("pagecount")
          If o_cfg("listtype") = "ul" Then
            i_listEnd = i_listEnd & " <li" & AddHtmlClass(0, o_cfg("pageclass")) & "><a href=""" & Replace(o_cfg("link"),"*",i) & """>" & Replace(o_cfg("list"),"*",i) & "</a></li> "
          Else
            i_listEnd = i_listEnd & " <a href=""" & Replace(o_cfg("link"),"*",i) & """" & AddHtmlClass(0, o_cfg("pageclass")) & ">" & Replace(o_cfg("list"),"*",i) & "</a> "
          End If
        Next
      End If
    End If
    '生成首页和上一页链接
    If o_cfg("pageindex") > 1 Then
      If o_cfg("listtype") = "ul" Then
        s_first = " <li" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("firstclass"))) & "><a href=""" & Replace(o_cfg("link"),"*","1") & """>" & o_cfg("first") & "</a></li> "
        s_prev = " <li" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("prevclass"))) & "><a href=""" & Replace(o_cfg("link"),"*",o_cfg("pageindex")-1) & """>" & o_cfg("prev") & "</a></li> "
      Else
        s_first = " <a href=""" & Replace(o_cfg("link"),"*","1") & """" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("firstclass"))) & ">" & o_cfg("first") & "</a> "
        s_prev = " <a href=""" & Replace(o_cfg("link"),"*",o_cfg("pageindex")-1) & """" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("prevclass"))) & ">" & o_cfg("prev") & "</a> "
      End If
    Else
      If o_cfg("listtype") = "ul" Then
        s_first = " <li class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("firstclass"))) & """><a href=""javascript:void(0)"">" & o_cfg("first") & "</a></li> "
        s_prev = " <li class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("prevclass"))) & """><a href=""javascript:void(0)"">" & o_cfg("prev") & "</a></li> "
      Else
        s_first = " <span class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("firstclass"))) & """>" & o_cfg("first") & "</span> "
        s_prev = " <span class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("prevclass"))) & """>" & o_cfg("prev") & "</span> "
      End If
    End If
    '生成下一页和末页链接
    If o_cfg("pageindex") < o_cfg("pagecount") Then
      If o_cfg("listtype") = "ul" Then
        s_next = " <li" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("nextclass"))) & "><a href=""" & Replace(o_cfg("link"),"*",o_cfg("pageindex")+1) & """>" & o_cfg("next") & "</a></li> "
        s_last = " <li" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("lastclass"))) & "><a href=""" & Replace(o_cfg("link"),"*",o_cfg("pagecount")) & """>" & o_cfg("last") & "</a></li> "
      Else
        s_next = " <a href=""" & Replace(o_cfg("link"),"*",o_cfg("pageindex")+1) & """" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("nextclass"))) & ">" & o_cfg("next") & "</a> "
        s_last = " <a href=""" & Replace(o_cfg("link"),"*",o_cfg("pagecount")) & """" & AddHtmlClass(0, Array(o_cfg("pageclass"), o_cfg("lastclass"))) & ">" & o_cfg("last") & "</a> "
      End If
    Else
      If o_cfg("listtype") = "ul" Then
        s_next = " <li class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("nextclass"))) & """><a href=""javascript:void(0)"">" & o_cfg("next") & "</a></li> "
        s_last = " <li class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("lastclass"))) & """><a href=""javascript:void(0)"">" & o_cfg("last") & "</a></li> "
      Else
        s_next = " <span class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("nextclass"))) & """>" & o_cfg("next") & "</span> "
        s_last = " <span class=""" & o_cfg("disabledclass") & AddHtmlClass(1, Array(o_cfg("pageclass"), o_cfg("lastclass"))) & """>" & o_cfg("last") & "</span> "
      End If
    End If
    Select Case LCase(o_cfg("jump"))
      Case "input"
        '生成跳转文本框
        s_jumpValue = "this.value"
        s_jump = " <input type=""text"" size=""3"" title=""" & Easp.Lang("db-pager-input-text") & """ " & Easp.IfHas(o_cfg("jumpplus"),"")
        '回车键执行跳转
        s_jump = s_jump & " onkeydown=""javascript:if(event.charCode==13||event.keyCode==13){if(!isNaN(" & s_jumpValue & ")){"
        s_jump = s_jump & Easp.IIF(o_cfg("jumpaction")="",Easp.IIF(Lcase(Left(o_cfg("link"),11))="javascript:",Replace(Mid(o_cfg("link"),12),"*",s_jumpValue),"document.location.href='" & Replace(o_cfg("link"),"*","'+" & s_jumpValue & "+'") & "';"),Replace(o_cfg("jumpaction"),"*", s_jumpValue))
        s_jump = s_jump & "}return false;}"" /> "
      Case "select"
        '生成跳转下拉框
        s_jumpValue = "this.options[this.selectedIndex].value"
        s_jump = " <select " & Easp.IfHas(o_cfg("jumpplus"),"") & " onchange=""javascript:"
        s_jump = s_jump & Easp.IIF(o_cfg("jumpaction")="",Easp.IIF(Lcase(Left(o_cfg("link"),11))="javascript:",Replace(Mid(o_cfg("link"),12),"*",s_jumpValue),"document.location.href='" & Replace(o_cfg("link"),"*","'+" & s_jumpValue & "+'") & "';"),Replace(o_cfg("jumpaction"),"*",s_jumpValue))
        s_jump = s_jump & """ title=""" & Easp.Lang("db-pager-select-text") & """> "
        '下拉框下拉菜单数量
        If o_cfg("jumplong")=0 Then
          For i = 1 To o_cfg("pagecount")
            s_jump = s_jump & " <option value=""" & i & """" & Easp.IfThen(i=o_cfg("pageindex")," selected=""selected""") & ">" & i & "</option> "
          Next
        Else
          i_jumpLong = Int(o_cfg("jumplong") / 2)
          i_jumpStart = Easp.IIF(o_cfg("pageindex")-i_jumpLong<1, 1, o_cfg("pageindex")-i_jumpLong)
          i_jumpStart = Easp.IIF(o_cfg("pagecount")-o_cfg("pageindex")<i_jumpLong, i_jumpStart-(i_jumpLong-(o_cfg("pagecount")-o_cfg("pageindex")))+1, i_jumpStart)
          i_jumpStart = Easp.IIF(i_jumpStart<1,1,i_jumpStart)
          j = 1
          For i = i_jumpStart To o_cfg("pageindex")
            s_jump = s_jump & " <option value=""" & i & """" & Easp.IfThen(i=o_cfg("pageindex")," selected=""selected""") & ">" & i & "</option> "
            j = j + 1
          Next
          i_jumpLong = Easp.IIF(o_cfg("pagecount")-o_cfg("pageindex")<i_jumpLong, i_jumpLong, i_jumpLong + (i_jumpLong-j)+1)
          i_jumpEnd = Easp.IIF(o_cfg("pageindex")+i_jumpLong>o_cfg("pagecount"), o_cfg("pagecount"), o_cfg("pageindex")+i_jumpLong)
          For i = o_cfg("pageindex")+1 To i_jumpEnd
            s_jump = s_jump & " <option value=""" & i & """>" & i & "</option> "
          Next
        End If
        s_jump = s_jump & "</select> "
    End Select
    '模板标签替换
    s_tmp = Replace(s_tmp,"{recordcount}",o_cfg("recordcount")) '总记录数
    s_tmp = Replace(s_tmp,"{pagecount}",o_cfg("pagecount")) '总页数
    s_tmp = Replace(s_tmp,"{pageindex}",o_cfg("pageindex")) '当前页码
    s_tmp = Replace(s_tmp,"{pagesize}",o_cfg("pagesize")) '每页条数
    s_tmp = Replace(s_tmp,"{rssize}",o_cfg("rssize")) '当前页记录数
    s_tmp = Replace(s_tmp,"{minrow}",o_cfg("minrow")) '当前页最小记录号
    s_tmp = Replace(s_tmp,"{maxrow}",o_cfg("maxrow")) '当前页最大记录号
    s_tmp = Replace(s_tmp,"{list}",s_list) '页码链接
    s_tmp = Replace(s_tmp,"{liststart}",i_listStart) '分页头部页码链接
    s_tmp = Replace(s_tmp,"{listend}",i_listEnd) '分页尾部页码链接
    s_tmp = Replace(s_tmp,"{first}",s_first) '首页链接
    s_tmp = Replace(s_tmp,"{prev}",s_prev) '上一页链接
    s_tmp = Replace(s_tmp,"{next}",s_next) '下一页链接
    s_tmp = Replace(s_tmp,"{last}",s_last) '末页链接
    s_tmp = Replace(s_tmp,"{jump}",s_jump) '页码跳转框
    s_tmp = "<" & o_cfg("listtype") & AddHtmlClass(0, o_cfg("listclass")) & ">" & s_tmp & "</" & o_cfg("listtype") & ">"
    Set o_cfg = Nothing
    Pager = s_tmp
  End Function
  '给html标签加上新的class
  Private Function AddHtmlClass(ByVal hasClass, ByVal nameArray)
    Dim i, s_tmp, b_tmp
    If Easp.IsN(nameArray) Then AddHtmlClass = "" : Exit Function
    If Not IsArray(nameArray) Then nameArray = Array(nameArray)
    For i = 0 To UBound(nameArray)
      If Easp.Has(nameArray(i)) Then
        If b_tmp Or hasClass = 1 Then s_tmp = s_tmp & " "
        s_tmp = s_tmp & nameArray(i)
        b_tmp = True
      End If
    Next
    If hasClass = 0 And Easp.Has(s_tmp) Then s_tmp = " class=""" & s_tmp & """"
    AddHtmlClass = s_tmp
  End Function

  '配置分页样式
  '参数：@pagerName    - 分页导航配置名称
  '     @html    - 分页导航模板
  '     @config  - 分页导航配置
  Public Sub SetPager(ByVal pagerName, ByVal html, ByRef config)
    pagerName = Easp.IfHas(pagerName, "default")
    If Easp.Has(html) Then o_pager(pagerName & "_html") = html
    If Easp.Has(config) Then o_pager(pagerName & "_config") = config
  End Sub

  '调用分页样式
  '参数：  @pagerName  - 分页导航配置名称
  '返回：  String
  Public Function GetPager(ByVal pagerName)
    pagerName = Easp.IfHas(pagerName, "default")
    GetPager = Pager(o_pager(pagerName & "_html"), o_pager(pagerName & "_config"))
  End Function

  '插入记录
  '参数：  @conn  - Connection对象
  '       @table - 数据表名
  '       @fieldValues - 字段名和值字符串，形如：
  '                       "field1:{value1}, field2:{post.value2:int}"
  '                       或者(省略字段名时需与数据表内字段数量、顺序完全一致)：
  '                       "{value1},{post.value2:int}"
  '返回：  numeric - 插入成功的记录数
  Public Function Insert(ByRef conn, ByVal table, ByVal fieldValues)
    On Error Resume Next
    Insert = InsertRecord(conn, table, fieldValues, False)
    CheckError "insert", Err, conn, "Insert", Array(table, fieldValues)
  End Function

  '用默认Connection插入记录
  Public Function Ins(ByVal table, ByVal fieldValues)
    On Error Resume Next
    OpenConn()
    Ins = InsertRecord(o_conn, table, fieldValues, False)
    CheckError "insert", Err, o_conn, "Ins", Array(table, fieldValues)
  End Function

  '批量插入记录
  Public Function InsertBatch(ByRef conn, ByVal table, ByVal fieldValues)
    On Error Resume Next
    InsertBatch = InsertRecord(conn, table, fieldValues, True)
    CheckError "insertbatch", Err, conn, "InsertBatch", Array(table, fieldValues)
  End Function

  '用默认Connection批量插入记录
  Public Function InsBatch(ByVal table, ByVal fieldValues)
    On Error Resume Next
    OpenConn()
    InsBatch = InsertRecord(o_conn, table, fieldValues, True)
    CheckError "insertbatch", Err, o_conn, "InsBatch", Array(table, fieldValues)
  End Function

  '取出参数中的每一对字段名称和值参数
  Private Function GetFiledValues(ByVal fieldValues, ByRef conn)
    Dim a_fieldValues, i_fieldValuesLength, i
    Dim hasFields, fields(), values()
    '按字段和值拆分
    a_fieldValues = Split(fieldValues, s_insSeparator)
    i_fieldValuesLength = UBound(a_fieldValues)
    '如果没有字段只有值
    hasFields = InStr(Easp.Str.Replace(fieldValues, "\{(.+?)\}", "?"), ":") > 0
    If hasFields Then ReDim fields(i_fieldValuesLength)
    ReDim values(i_fieldValuesLength)
    '取出字段名称和值参数
    For i = 0 To i_fieldValuesLength
      If hasFields Then
        fields(i) = FixName(Trim(Easp.Str.GetColonName(a_fieldValues(i))), conn)
        values(i) = Trim(Easp.Str.GetColonValue(a_fieldValues(i)))
      Else
        '如果没有字段名
        values(i) = Trim(a_fieldValues(i))
      End If
    Next
    GetFiledValues = Array(hasFields, Fields, values)
  End Function

  '插入记录原型
  Private Function InsertRecord(ByRef conn, ByVal table, ByVal fieldValues, ByVal IsBatch)
    Dim a_fieldValues, i, a_tmp, s_dbType, a_values
    Dim s_sqlstart, s_sqlvalues, s_sqlend, i_result, i_paramCount
    Dim i_sqllimit, i_paramlimit, i_limit
    Dim i_valuesCount, i1, i2, i3, i4, s_sqltmp
    table = FixName(table, conn)
    '取数据库类型
    s_dbType = GetType(conn)
    '拆分字段和值
    a_fieldValues = GetFiledValues(fieldValues, conn)
    '组合为SQL插入语句
    If a_fieldValues(0) Then
      s_sqlstart = "Insert Into " & table & " (" & Join(a_fieldValues(1),", ") & ") Values "
    Else
      s_sqlstart = "Insert Into " & table & " Values "
    End If
    s_sqlvalues = "(" & Join(a_fieldValues(2), ", ") & ")"
    '构造批量插入语句
    If IsBatch Then
      On Error Resume Next '## Do NOT delete or comment this line
      Err.Clear
      i_result = 0
      '开始事务
      BeginTrans(conn)
      a_tmp = GetBatchArray(s_sqlvalues)
      a_values = a_tmp(0)
      i_paramCount = a_tmp(1)
      i_valuesCount = UBound(a_values)+1
      'MSSQL数据库
      If s_dbType = "MSSQL" Or s_dbType = "MYSQL" Then
        Select Case s_dbType
          Case "MSSQL"
            'MSSQL参数不能超过2100个，同时插入不能超过1000条
            i_paramlimit = 2000
            i_sqllimit = 1000
            i_sqllimit = Easp.IIF(Int(i_paramlimit/i_paramCount)>i_sqllimit,i_sqllimit,Int(i_paramlimit/i_paramCount))
          Case "MYSQL"
            'MySQL数据库限制为同时最多插入5000条，避免存储空间不足错误
            i_sqllimit = 5000
        End Select
        i_limit = i_valuesCount/i_sqllimit
        i_limit = Int(i_limit) + Easp.IIF(i_limit>Int(i_limit),1,0) - 1
        For i1 = 0 To i_limit
          i2 = i1*i_sqllimit
          i3 = i1*i_sqllimit + i_sqllimit - 1
          If i3>(i_valuesCount-1) Then i3 = (i_valuesCount-1)
          s_sqltmp = ""
          For i4 = i2 To i3
            If i4 > i2 Then s_sqltmp = s_sqltmp & ","
            s_sqltmp = s_sqltmp & a_values(i4)
          Next
          '执行批量插入
          'Easp.console i_limit & " / " & i1 & " / " & i2 & " / " & i3 & " / " & i4
          i_result = i_result + ExecuteSql(conn, s_sqlstart & s_sqltmp, 0)
        Next
      Else
        'Access只能一条条插入
        For i = 0 To UBound(a_values)
          i_result = i_result + ExecuteSql(conn, s_sqlstart & a_values(i), 0)
        Next
      End If
      If Err.number = 0 And conn.Errors.Count = 0 Then
        '提交事务
        CommitTrans(conn)
      Else
        RollbackTrans(conn)
        'CheckError "insertbatch", Err, conn, "InsertBatch / InsBatch", Array(table, fieldValues)
      End If
      InsertRecord = i_result
      On Error GoTo 0
    Else
      InsertRecord = ExecuteSql(conn, s_sqlstart & s_sqlvalues, 0)
    End If
  End Function

  '删除记录
  Public Function Delete(ByRef conn, ByVal table, ByVal where)
    On Error Resume Next
    Dim sql
    table = FixName(table, conn)
    sql = "Delete From " & table
    If Easp.Has(where) Then sql = sql & " Where " & where
    Delete = ExecuteSql(conn, sql, 0)
    CheckError "delete", Err, conn, "Delete/Del", sql
  End Function
  '用默认Connection删除记录
  Public Function Del(ByVal table, ByVal where)
    OpenConn()
    Del = Delete(o_conn, table, where)
  End Function
  '批量删除记录
  Public Function DeleteBatch(ByRef conn, ByVal table, ByVal where)
    On Error Resume Next
    Dim sql
    table = FixName(table, conn)
    sql = "Delete From " & table
    If Easp.Has(where) Then
      sql = sql & " Where ("
      sql = sql & Join(GetBatchArray(where)(0),") Or (")
      sql = sql & ")"
    End If
    DeleteBatch = ExecuteSql(conn, sql, 0)
    CheckError "deletebatch", Err, conn, "DeleteBatch/DelBatch", sql
  End Function
  '用默认Connection批量删除记录
  Public Function DelBatch(ByVal table, ByVal where)
    OpenConn()
    DelBatch = DeleteBatch(o_conn, table, where)
  End Function

  '更新记录
  Public Function Update(ByRef conn, ByVal table, ByVal fieldValues, ByVal where)
    On Error Resume Next
    Dim sql
    table = FixName(table, conn)
    sql = "Update " & table & " Set " & fieldValues
    If Easp.Has(where) Then sql = sql & " Where " & where
    Update = ExecuteSql(conn, sql, 0)
    CheckError "update", Err, conn, "Update/Upd", sql
  End Function
  '用默认Connection删除记录
  Public Function Upd(ByVal table, ByVal fieldValues, ByVal where)
    OpenConn()
    Upd = Update(o_conn, table, fieldValues, where)
  End Function
  '批量删除记录
  Public Function UpdateBatch(ByRef conn, ByVal table, ByVal fieldValues, ByVal where)
    On Error Resume Next
    Dim sql, a_field
    table = FixName(table, conn)
    a_field = GetBatchArray(fieldValues)(0)
    sql = "Update " & table & " Set " & fieldValues
    If Ubound(a_field) > 0 Then
      If Easp.Has(where) Then sql = sql & " Where " & where
      UpdateBatch = ExecuteBatch(conn, sql)
    Else
      If Easp.Has(where) Then
        sql = sql & " Where (" & Join(GetBatchArray(where)(0), ") Or (") & ")"
      End If
      UpdateBatch = ExecuteSql(conn, sql, 0)
    End If
    CheckError "updatebatch", Err, conn, "UpdateBatch/UpdBatch", sql
  End Function
  '用默认Connection批量删除记录
  Public Function UpdBatch(ByVal table, ByVal fieldValues, ByVal where)
    OpenConn()
    UpdBatch = UpdateBatch(o_conn, table, fieldValues, where)
  End Function

  '表名关键字处理
  Private Function FixName(ByVal string, ByRef conn)
    string = Trim(string)
    Select Case UCase(GetType(conn))
      Case "ACCESS", "MSSQL"
        If Not Easp.Str.Test(string, "^(\[.+\]|"".+"")$") Then
          string = "[" & string & "]"
        End If
      Case "MYSQL"
        If Not Easp.Str.Test(string, "^`.+`$") Then
          string = "`" & string & "`"
        End If
      Case Else
        string = """" & straing & """"
    End Select
    FixName = string
  End Function
  '替换SQL语句中的{easp.newid}
  Private Function ReplaceNewId(ByVal sql)
    If InStr(1,sql, "{easp.newid}",1)>0 Then
      ReplaceNewId = Replace(sql, "{easp.newid}", "'" & Easp.NewID() & "'", 1, -1, 1)
    Else
      ReplaceNewId = sql
    End If
  End Function

  '替换SQL语句中的静态变量 {=Easp变量名}
  Private Function ReplaceStasicParameter(ByVal sql)
    Dim matches, Match
    If Instr(sql, "{=") Then
      Set matches = Easp.Str.Match(sql, "\{=(.+?)\}")
      For Each match In matches
        'Easp.Consoel match
        sql = Replace(sql, match, Easp.Var(match.SubMatches(0)), 1, -1, 1)
      Next
      Set matches = Nothing
    End If
    ReplaceStasicParameter = sql
  End Function

  '转换sql参数类型为数值
  Private Function GetParameterType(ByVal paramType)
    Dim n
    If IsNumeric(paramType) Then
      n = paramType
    Else
      Select Case LCase(paramType)
        'from the ASP book
        Case "empty"              n = 0
        Case "smallint"           n = 2
        Case "integer"            n = 3
        Case "single"             n = 4
        Case "double"             n = 5
        Case "currency"           n = 6
        Case "date"               n = 7
        Case "bstr"               n = 8
        Case "idispatch"          n = 9
        Case "error"              n = 10
        Case "boolean"            n = 11
        Case "variant"            n = 12
        Case "iunknown"           n = 13
        Case "decimal"            n = 14
        Case "tinyint"            n = 16
        Case "unsignedtinyint"    n = 17
        Case "unsignedsmallint"   n = 18
        Case "unsignedint"        n = 19
        Case "bigint"             n = 20
        Case "ebigint"            n = 20
        Case "unsignedbigint"     n = 21
        Case "guid"               n = 72
        Case "binary"             n = 128
        Case "char"               n = 129
        Case "wchar"              n = 130
        Case "numeric"            n = 131
        Case "userdefined"        n = 132
        Case "dbdate"             n = 133
        Case "dbtime"             n = 134
        Case "dbtimestamp"        n = 135
        Case "varchar"            n = 200
        Case "longchar"           n = 201
        Case "longvarchar"        n = 201
        Case "memo"               n = 201
        Case "varwchar"           n = 202
        Case "longvarwchar"       n = 203
        Case "string"             n = 201
        Case "varbinary"          n = 204
        Case "longvarbinary"      n = 204
        Case "longbinary"         n = 205
        'SQL Server
        Case "bit"                n = 11
        Case "money"              n = 6
        Case "int"                n = 3
        Case "smallmoney"         n = 6
        Case "float"              n = 5
        Case "nchar"              n = 200
        Case "real"               n = 131
        Case "text"               n = 200
        Case "time"               n = 134
        Case "timestamp"          n = 135
        Case "datetime"           n = 135
        Case "smalldatetime"      n = 135
        Case "datetime2"          n = 135
        Case "sysname"            n = 129
        Case "uniqueidentifier"   n = 131
        Case "ntext"              n = 200
        Case "nvarchar"           n = 200
        Case "nvarchar2"          n = 200
        Case "image"              n = 204
        Case "sql_variant"        n = 12
        'MySQL
        Case "year"               n = 133
        Case "tinytext"           n = 200
        Case "mediumtext"         n = 200
        Case "longtext"           n = 201
        Case "mediumint"          n = 3
        Case "enum"               n = 132
        Case "set"                n = 132
        'others
        Case "byte"               n = 16
        Case "long"               n = 3
        Case "counter"            n = 131
        Case Else                 n = 200
      End Select
    End If
    GetParameterType = n
  End Function

  '取参数名称、类型和值
  Private Function GetParameter(ByVal param)
    Dim arr(2)
    param = Trim(param)
    If Left(param,1)<>"{" Or Right(param,1)<>"}" Then Exit Function
    param = Mid(param, 2, Len(param)-2)
    arr(0) = Easp.Str.GetColonName(param)
    arr(1) = Easp.Str.GetColonValue(param)
    arr(2) = Easp.Var(arr(0))
    GetParameter = arr
  End Function

  '将值转换为输出SQL时可显示的数据
  Private Function FormatValue(ByVal value, ByVal ptype)
    Dim s_tmp
    Select Case ptype
      Case 129,130,200,201,202,203
        If Easp.IsN(value) Then
          s_tmp = "''"
        Else
          s_tmp = "'" & Replace(value, "'", "''") & "'"
        End If
      Case 7,133,134,135
        If Easp.IsN(value) Then
          s_tmp = "NULL"
        Else
          s_tmp = "'" & Replace(value, "'", "''") & "'"
        End If
      Case 13, 128, 204, 205
        s_tmp = Easp.IIF(Easp.Has(value),"(blob)","NULL")
      Case Else
        s_tmp = Easp.IfHas(value,"NULL")
        s_tmp = Replace(s_tmp, "'", "''")
    End Select
    FormatValue = s_tmp
  End Function
  
  '关闭并释放对象
  '参数：  @obj  - ASP对象
  '返回：  无
  Public Sub Close(ByRef obj)
    If TypeName(obj) = "Recordset" Or TypeName(obj) = "Connection" Then
      If obj.State <> 0 Then obj.close()
    End If
    Set obj = Nothing
  End Sub
  
  '开始一个事务
  Public Function BeginTrans(ByRef conn)
    BeginTrans = conn.BeginTrans
    If Easp.Console.ShowSql Then Easp.Console Easp.Str.Format(Easp.Lang("db-trans-start"), BeginTrans)
  End Function
  '回滚一个事务
  Public Sub RollbackTrans(ByRef conn)
    conn.RollbackTrans
    If Easp.Console.ShowSql Then Easp.Console Easp.Lang("db-trans-rollback")
  End Sub
  '提交一个事务
  Public Sub CommitTrans(ByRef conn)
    conn.CommitTrans
    If Easp.Console.ShowSql Then Easp.Console Easp.Lang("db-trans-commit")
  End Sub
  '开始默认连接的事务
  Public Sub Begin()
    OpenConn()
    i_transLevel = BeginTrans(o_conn)
  End Sub
  '回滚默认连接的事务
  Public Sub Rollback()
    OpenConn()
    RollbackTrans(o_conn)
    i_transLevel = i_transLevel - 1
  End Sub
  '提交默认连接的事务
  Public Sub Commit()
    OpenConn()
    CommitTrans(o_conn)
    i_transLevel = i_transLevel - 1
  End Sub
End Class
%>