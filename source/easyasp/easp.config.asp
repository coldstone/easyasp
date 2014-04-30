<%
'######################################################################
'## easp.config.asp
'## -------------------------------------------------------------------
'## EasyASP 配置文件
'######################################################################

'加载语言包文件
%>
<!--#include file="lang/zh.asp" -->
<%
Easp.Init() 'EasyASP初始化
''====================================
''  EasyASP 基础配置
''====================================

'设置'easp.asp'文件在网站中的路径，以"/"开头:
'Easp.BasePath = "/easyasp/"
'设置EasyASP存放插件文件的路径，以"/"开头:
'Easp.PluginPath = "/easyasp/plugin"
'打开开发者调试模式：
Easp.Debug = True
'配置默认首页的文件名，用于伪Rewrite时的首页地址省略
'Easp.DefaultPageName = "index.asp"

''====================================
''  Db 数据库配置
''====================================

'配置数据库默认连接：
'Easp.Db.SetConn "ACCESS", "/sampledata/EaspSampleData.mdb", ""
'Easp.Db.SetConn "MSSQL", "data", "sa:pass@(local)"
'配置第二个数据库连接
'Easp.Db.SetConnection "connname", "MSSQL", "data", "sa:pass@(local)"
'Easp.Db.SetConnection "connname", 1, "data", "root:pass@server:port"
'设置分页标识URL参数
'Easp.Db.PageParam          = "page"
'设置分页每页数量
'Easp.Db.PageSize           = 25

''====================================
''  Console 控制台配置
''====================================

'在这里设置token的值，区分大小写，如果设置了Token值，
'仅前端输入的token和这里设置一致时，才会输出控制台信息
Easp.Console.Token          = ""
''是否开启控制台
Easp.Console.Enable         = True
''是否在控制台中自动显示执行的SQL语句
'Easp.Console.ShowSql       = True
''是否在控制台中自动显示执行的SQL语句的执行时间
'Easp.Console.ShowSqlTime   = True
''控制台中缓存的内容最大字节数
'Easp.Console.MaxCacheSize  = 8000
''单条控制台输出的内容最大字节数
'Easp.Console.MaxLogSize    = 3000

''====================================
''  Error 异常信息配置
''====================================
'抛出异常信息时的标题
'Easp.Error.Title           = "发生错误啦"
'是否自动跳转页面
'Easp.Error.Redirect        = True
'跳转等待时间（秒）
'Easp.Error.Delay           = 5
'错误信息框的css样式
'Easp.Error.ClassName       = ""

''====================================
''  Fso 配置
''====================================

'设置FSO组件名称（如果服务器上修改过）
'Easp.Fso.fsoName           = "Scripting.FileSystemObject"
'设置是否删除只读文件
'Easp.Fso.Force             = True
'设置是否覆盖原有文件
'Easp.Fso.OverWrite         = True
'设置文件大小显示格式(G,M,K,b,auto)
'Easp.Fso.SizeFormat        = "K"
'设置UTF-8文件的BOM信息如何处理(keep/remove/add)
'Easp.Fso.FileBom           = "keep"

''====================================
''  Http 配置
''====================================

'异步模式
'Easp.Http.Async            = False
'服务器解析超时（毫秒）
'Easp.Http.ResolveTimeout   = 20000
'服务器连接超时（毫秒）
'Easp.Http.ConnectTimeout   = 20000
'发送数据超时（毫秒）
'Easp.Http.SendTimeout      = 300000
'接受数据超时（毫秒）
'Easp.Http.ReceiveTimeout   = 60000

''====================================
''  Json 配置
''====================================

'设置生成Json字符串是是否编码 Unicode 字符
'Easp.Json.EncodeUnicode    = True

''====================================
''  List 配置
''====================================

'List的键值是否区分大小写
'Easp.List.IgnoreCase      = False

''====================================
''  Str 配置
''====================================

'是否编码ToString时的Unicode字符
'Easp.Str.EncodeJsonUnicode = False

''====================================
''  Tpl 配置
''====================================

'设置静态模板文件存放的目录路径
'Easp.Tpl.FilePath          = "/view/"
'设置和读取标签的标识符
'Easp.Tpl.TagMask           = "{*}"
'设置模板中是否可以执行ASP代码
'Easp.Tpl.AspEnable         = False
'设置如何处理未定义的标签(keep/remove/comment)
'Easp.Tpl.TagUnknown        = "keep"

''====================================
''  Upload 配置
''====================================

'配置文本编码
'Easp.Upload.CharSet          = "utf-8"
'设置允许上传的最大字节数
'Easp.Upload.AllowMaxSize     = -1
'设置允许上传的单文件的最大字节数
'Easp.Upload.AllowMaxFileSize = -1
'设置允许上传的文件的扩展名
'Easp.Upload.AllowFileTypes   = "jpg|png|gif|.rar|.zip|doc|docx|xls|xlsx|ppt|pptx|pdf"
'设置文件保存目录
'Easp.Upload.SavePath         = "/UploadFiles"
%>