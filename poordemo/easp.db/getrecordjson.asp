<!--#include file="_cls.testcase.asp" --><%
Dim rs
Set rs = Easp.Db.Query("Select ContentID As id, ContentClassID As cid, ContentTitle As title, AnnounceTime As atime, ContentText As content From EC_Content Where ContentID = {id}")
Easp.Print Easp.Encode(rs)
Easp.Db.Close(rs)
%>