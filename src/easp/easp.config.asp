<%
'######################################################################
'## easp.config.asp
'## -------------------------------------------------------------------
'## EasyASP 配置文件
'######################################################################

'必须正确设置'easp.asp'文件在网站中的路径，以"/"开头:
Easp.BasePath = "/easp/"

'设置文件编码 (通常为'GBK'或者'UTF-8'):
Easp.CharSet = "UTF-8"

''打开开发者调试模式：
Easp.Debug = False

''不加密Cookies数据:
'Easp.CookieEncode = False

''设置FSO组件的名称（如果服务器上修改过）:
'Easp.FsoName = "Scripting.FileSystemObject"

''设置如何处理载入的UTF-8文件的BOM信息(keep/remove/add)：
'Easp.FileBOM = "remove"

''配置数据库连接：
''Access:
'Easp.db.Conn = Easp.db.OpenConn(1,"/data/data.mdb","")
''MS SQL Server:
'Easp.db.Conn = Easp.db.OpenConn(0,"Data","sa:admin@127.0.0.1")
%>