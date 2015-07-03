<!DOCTYPE html>
<html>
<head>
 <meta charset="utf-8" />
 <title>HTML多文件批量上传</title>
</head>
<body>
多文件批量上传<br />
<form action="upload.asp" method="post" enctype="multipart/form-data">
文件1：<input type="file" id="file1" name="file1" multiple="multiple" /><br />
文件2：<input type="file" id="file2" name="file2" multiple="multiple" /><br />
文件3：<input type="file" id="file3" name="file3" multiple="multiple" /><br /><input type="submit" value="上传" />
</form>
</body></html>