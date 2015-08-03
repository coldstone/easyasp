<!--#include file="../../easyasp/easp.asp" --><%
'所有日志文件默认保存在站点根目录同一父目录下，站点文件夹名称后加_log的文件夹内
'如果保存失败，请确认是否有写入权限
Easp.Log.Enable = True
Easp.Log.Style("info") = Easp.Log.Style("info") & ", {note}, {param}"
Easp.Log.Set "note", "所有的信息中都有"
Easp.Log.SetOne "param", "只替换一次"
Easp.Log.Info "来测试一条信息吧"
Easp.Log.Info "再来测试一条，和上面不同哦"
Easp.Log.Warn "使用默认的模板输出警告信息"
Easp.Log.Error "这里出错啦", "问题出在这个文件(定位):index.asp:9"

%>