<!--#include file="_cls.testcase.asp" -->
<%
Dim result, act
act = LCase(Easp.Get("action"))
Select Case act
  Case "savenew", "saveedit"  '添加或编辑内容
    '在这里编写你的服务端表单验证代码
    Call Easp.PostVal("title").Name("标题").Required.Alert
    Call Easp.PostVal("classid").Name("所属类别").Required.Maxlength(10).Test("^\w+$").Msg("%n不是正确的ID序号").Alert
    Easp.Var("post.announcetime") = Easp.VarVal("post.announcetime").Name("发布时间").Default(Now()).IsDate.Alert
    Easp.PostVal("content").Name("内容").Required.Alert
    '此处仅作简单的空值判断
    'If Easp.IsN(Easp.Var("post.title")) Or Easp.IsN(Easp.Var("post.classid")) Or Easp.IsN(Easp.Var("post.announcetime")) Or Easp.IsN(Easp.Var("post.content")) Then
    '  Easp.RR "?info=" & act & "error"
    'End If
    If act = "savenew" Then
      '添加到数据库
      result = Easp.Db.Ins("EC_Content", "ContentID:{easp.newid}, ContentClassID:{post.classid}, ContentTitle:{title}, AnnounceTime:{announcetime},ContentText:{content}")
    Else
      result = Easp.Db.Upd("EC_Content", "ContentClassID={post.classid}, ContentTitle={title}, AnnounceTime={announcetime},ContentText={content}", "ContentID={cid}")
    End If
    If result Then Easp.RR "?info=" & act & "success"
  Case "delete"
    '判断有没有选择内容
    If Easp.IsN(Easp.Var("id")) Then Easp.RR "?info=deleteempty"
    '批量删除（post和get方法提交的数据均可）
    result = Easp.Db.DelBatch("EC_Content","ContentID = {id}")
    Easp.Console "delete " & result & " records."
    Easp.RR "?info=deletesuccess"
End Select

'取内容类别
Function GetContentClass(ByVal default)
  GetContentClass = TestCase.Db.GetSelectOptions("Select  ContentClassID, ContentClassName, ParentID From EC_ContentClass Where IsDeleted = False Order By SortLevel Asc", default, "-1")
End Function
'显示信息提示栏
Function ShowAlert()
  Dim str,arr
  str = "<div class=""alert alert-{0} alert-dismissable"">" &_
        "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-hidden=""true"">&times;</button>" &_
        "<strong>{1}</strong> {2}</div>"
  Select Case Easp.Var("info")
    Case "savenewsuccess"
      arr = Array("success", "添加新内容成功！","")
    Case "savenewerror"
      arr = Array("danger", "添加新内容出错！","请完整输入必填项。")
    Case "saveeditsuccess"
      arr = Array("success", "编辑内容成功！","")
    Case "saveediterror"
      arr = Array("danger", "编辑内容出错！","请完整输入必填项。")
    Case "deletesuccess"
      arr = Array("success", "删除内容成功！","")
    Case "deleteempty"
      arr = Array("danger", "删除内容出错！","请先选择要删除的内容。")
    Case Else
      ShowAlert = ""
      Exit Function
  End Select
  ShowAlert = Easp.Str.Format(str,arr)
End Function
%>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Access数据库增删改查示例 - EasyASP测试用例中心</title>
    <!-- Bootstrap -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.1.0/css/bootstrap.min.css">
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.1.0/css/bootstrap-theme.min.css">
    <link rel="stylesheet" href="style.css">
    <!--[if lt IE 9]>
      <script src="http://cdn.bootcss.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="http://cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>
  <body>

<!--#include file="_nav.asp" -->

    <div class="container">

      <div class="row row-offcanvas row-offcanvas-right">

        <div class="col-xs-6 col-sm-3 sidebar-offcanvas" id="sidebar" role="navigation">
          <div class="list-group">
            <a href="index.asp" class="list-group-item active">Access数据库增删改查</a>
            <a href="access.insertbatch.asp" class="list-group-item">Access数据库批量添加</a>
          </div>
        </div><!--/span-->

        <div class="col-xs-12 col-sm-9">
          <%=ShowAlert()%>
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">内容列表</h3>
            </div>
            <div class="panel-body">
              <div class="row m-b-10">
                <div class="pull-left p-l-20">
                  <button type="button" class="btn btn-default btn-sm" data-toggle="modal" data-target="#win-new-info" id="btn-new-info"><span class="glyphicon glyphicon-plus"></span> 添加</button>
                  <button type="button" class="btn btn-default btn-sm" id="btn-del-batch"><span class="glyphicon glyphicon-minus"></span> 删除所选</button>
                </div>
                <form class="form-inline pull-left p-l-10" role="searchtitle" action="?" method="get">
                  <div class="form-group">
                    <input type="text" class="form-control input-sm" placeholder="搜索标题..." name="key">
                  </div>
                  <div class="form-group btn-group btn-group-sm">
                    <button type="submit" class="btn btn-default btn-sm"><span class="glyphicon glyphicon-search"></span> 搜索</button>
                    <button type="button" class="btn btn-default btn-sm dropdown-toggle" data-toggle="dropdown">
                      <span class="caret"></span>
                    </button>
                    <ul class="dropdown-menu" role="menu">
                      <li><a class="small" href="?">重置</a></li>
                    </ul>
                  </div><!-- /btn-group -->
                </form>
              </div><!-- /row -->
              <table class="table table-hover table-condensed">
                <thead>
                  <th>&nbsp;</th>
                  <th><a href="?">全部类别</a></th>
                  <th>标题</th>
                  <th>发布时间</th>
                  <th>操作</th>
                </thead><form action="?action=delete" method="post" id="formdel"><%
Dim rs
'这里演示下如何使用使用Like需要的参数，通过在超级变量中嵌入静态标签就组合成了新的参数
Easp.Var("likeKey") = "%{=key}%"
'获取内容记录集
Set rs = Easp.Db.GetRS("Select a.ContentID, a.ContentClassID, b.ContentClassName, a.ContentTitle, a.AnnounceTime, b.SortLevel From EC_Content a Inner join EC_ContentClass b On a.ContentClassID = b.ContentClassID Where a.AnnounceTime<=NOW() And ({classid}='' Or a.ContentClassID = {classid}) And ({key}='' Or a.ContentTitle Like {likeKey}) Order By AnnounceTime Desc")
If Easp.Has(rs) Then
  'Easp.Console rs
  While Not rs.Eof
                %>
                <tr>
                  <td><input type="checkbox" name="id" value="<%=rs("ContentID")%>"></td>
                  <td><%= Easp.Str.Format("<a href=""?classid={ContentClassID}"">[ {ContentClassName:EHtmlEncode(%s)} ]</a>", rs)%></td>
                  <td><%= Easp.Str.HtmlEncode(rs("ContentTitle"))%></td>
                  <td><%= Easp.Date.Format(rs("AnnounceTime"), "y-mm-dd")%></td>
                  <td>
                    <button type="button" class="btn btn-primary btn-sm btn-edit-info" data-toggle="modal" data-target="#win-new-info" infoid="<%=rs("ContentID")%>"><span class="glyphicon glyphicon-edit"></span> 编辑</button>
                    <a href="?action=delete&id=<%=rs("ContentID")%>" class="btn btn-primary btn-sm"><span class="glyphicon glyphicon-minus-sign"></span> 删除</a>
                  </td>
                </tr><%
    rs.MoveNext()
  Wend
Else
                %>
                <tr>
                  <td colspan="5">没有数据</td>
                </tr><%
End If
Easp.Db.Close(rs)
                %></form>
              </table>
              <%=Easp.Db.GetPager("bootstrap")%>
              <div class="modal fade" id="win-new-info" role="dialog" aria-hidden="true" tabindex="-1">
                <div class="modal-dialog">
                  <div class="modal-content">
                    <div class="modal-header">
                      <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                      <h4 class="modal-title" id="modal-title">添加新内容</h4>
                    </div>
                    <div class="modal-body">
                      <form action="?action=savenew" method="post" class="form" id="form-addnew">
                        <input type="hidden" name="cid" id="cid" value="">
                        <div class="form-group">
                          <label for="title" class="control-label">标题：</label>
                          <input type="text" name="title" id="title" class="form-control" value="" placeholder="请输入标题..." />
                        </div>
                        <div class="form-group">
                          <label for="title" class="control-label">所属类别：</label>
                          <select class="form-control" name="classid" id="classid"><%=GetContentClass("")%></select>
                        </div>
                        <div class="form-group">
                          <label for="announcetime" class="control-label">发布时间：</label>
                          <input type="text" name="announcetime" id="announcetime" class="form-control" value="<%=Now%>" />
                        </div>
                        <div class="form-group">
                          <label for="content" class="control-label">内容：</label>
                          <textarea name="content" id="content" class="form-control" rows="5" placeholder="请输入内容..."></textarea>
                        </div>
                      </form>
                    </div>
                    <div class="modal-footer">
                      <button type="button" class="btn btn-default" data-dismiss="modal">取消</button>
                      <button type="button" class="btn btn-primary" id="btn-sumbit">保存</button>
                    </div>
                  </div><!-- /.modal-content -->
                </div><!-- /.modal-dialog -->
              </div><!-- /.modal add -->

            </div><!-- /panel-body -->
          </div>
        </div><!--/span-->

      </div><!--/row-->

      <hr>

      <footer>
        <div class="row">
          <p class="col-sm-4">&copy; EasyASP v3, 2014</p>
          <p class="col-sm-8"><span class="pull-right"><%=TestCase.ShowScriptTime()%></span></p>
        </div>
      </footer>

    </div><!--/.container-->

    <script src="//cdn.bootcss.com/jquery/1.11.1/jquery.min.js"></script>
    <script src="//cdn.bootcss.com/bootstrap/3.1.0/js/bootstrap.min.js"></script>
    <!--<script src="//netdna.bootstrapcdn.com/bootstrap/3.1.0/js/bootstrap.min.js"></script>-->
    <script type="text/javascript">
      $(document).ready(function(){
        $("#btn-sumbit").on("click",function(){
          if(Easp.validation("#title,#classid,#announcetime,#content"))
            $("#form-addnew").submit();
        });
        $("#btn-del-batch").on("click",function(){
          $("#formdel").submit();
        });
        $("#btn-new-info").on("click",function(){
          $("#modal-title").text("添加新内容");
          Easp.formclear();
          $("#form-addnew").attr("action", "?action=savenew");
        });
        $(".btn-edit-info").on("click",function(){
          $("#modal-title").text("编辑内容");
          Easp.edit($(this).attr('infoid'));
        });
      });
      var Easp = window.Easp = {
        validation : function(el){
          var flag = true;
          $(el).each(function(){
            var $el = $(this);
            //alert($el.val());
            if($el.val()){
              $el.parent().removeClass("has-error").addClass("has-success");
            } else {
              $el.parent().removeClass("has-success").addClass("has-error");
              flag = false;
            }
          });
          return flag;
        },
        formclear : function(){
          $("#title,#content,#classid").val("");
        },
        edit : function(id){
          Easp.formclear();
          $("#form-addnew").attr("action", "?action=saveedit");
          $.get("getrecordjson.asp", {id:id,random:Math.random()}, function(data){
            var _data = eval("(" + data + ")");
            if (_data.total>0){
              var row = _data.rows[0];
              $("#cid").val(row.id);
              $("#title").val(row.title);
              $("#classid").val(row.cid);
              $("#announcetime").val(row.atime);
              $("#content").val(row.content);
            }
          });
        }
      };
    </script>
  </body>
</html>