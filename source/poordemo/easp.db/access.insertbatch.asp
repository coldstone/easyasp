<!--#include file="_cls.testcase.asp" -->
<%
'Easp.Console Easp.Db.GetConn.Provider
'Easp.Console.ShowSqlTime = False
Dim result
Select Case LCase(Easp.Var("action"))
  Case "savenew"  '添加新内容
    '在这里编写你的服务端表单验证代码（此实例省略）
    '此处仅作简单的空值判断
    If Easp.IsN(Easp.Var("post.classid")) Or Easp.IsN(Easp.Var("post.name")) Then
      Easp.RR "?info=adderror"
    End If
    '添加到数据库
    result = Easp.Db.InsBatch("EC_ContentClass", "ContentClassID:{easp.newId}, ContentClassName:{name}, ParentID:{classid}, SortLevel:{sort:int}")
    If result>0 Then Easp.RR "?info=addsuccess"
End Select

'取内容类别
Function GetContentClass()
  GetContentClass = TestCase.Db.GetSelectOptions("Select ContentClassID, ContentClassName, ParentID From EC_ContentClass Where IsDeleted = False Order By SortLevel Asc", "", "-1")
End Function
'显示信息提示栏
Function ShowAlert()
  Dim str,arr
  str = "<div class=""alert alert-{0} alert-dismissable"">" &_
        "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-hidden=""true"">&times;</button>" &_
        "<strong>{1}</strong> {2}</div>"
  Select Case Easp.Var("info")
    Case "addsuccess"
      arr = Array("success", "添加成功！","")
    Case "adderror"
      arr = Array("danger", "添加出错！","请完整输入必填项。")
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
    <title>Access数据库批量添加示例 - EasyASP测试用例中心</title>
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
            <a href="index.asp" class="list-group-item">Access数据库增删改查</a>
            <a href="access.insertbatch.asp" class="list-group-item active">Access数据库批量添加</a>
          </div>
        </div><!--/span-->

        <div class="col-xs-12 col-sm-9">
          <%=ShowAlert()%>
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">批量添加类别</h3>
            </div>
            <div class="panel-body">

              <form action="?action=savenew" method="post" class="form" id="form-addnew">
                <div class="form-group form-inline">
                  <label for="title" class="control-label">所属父类：</label>
                  <select class="form-control" name="classid" id="classid"><%=GetContentClass()%><option value="-1">+&lt;新的大类&gt;</option></select>
                </div>
                <div class="form-group form-inline" id="row-name">
                  <label for="name1" class="control-label">类别名称：</label>
                  <input type="text" name="name" id="name1" class="form-control" placeholder="请输入类别名称..." />
                  <label for="sort" class="control-label">排序号：</label>
                  <input type="text" name="sort" id="sort" class="form-control" placeholder="999" />
                  <button type="button" name="additem" class="btn btn-default btn-sm"><span class="glyphicon glyphicon-plus"></span> 新增一行</button>
                  <button type="button" name="delitem" class="btn btn-default btn-sm hide" onclick="del(this)"><span class="glyphicon glyphicon-minus"></span> 删除此行</button>
                </div>
                <div class="form-group form-inline" id="submitbtn">
                  <button type="submit" class="btn btn-default"><span class="glyphicon glyphicon-floppy-saved"></span> 保存</button>
                </div>
              </form>

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
    <script type="text/javascript">
      $(document).ready(function(){
        $("button[name=additem]").on("click",function(){
          $("#row-name").clone()
            .insertBefore("#submitbtn").removeAttr("id")
            .find("button[name=additem]").hide()
              .next().removeClass("hide");
        });
      });
      function del(el){
        $(el).parent().remove();
      }
    </script>
  </body>
</html>