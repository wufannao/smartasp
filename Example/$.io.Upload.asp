<!--#include file="../Src/SmartASP.Core.asp"-->
<!--#include file="../Src/SmartASP.IO.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>文件上传</title>
</head>

<body>
<form action="?action=post" method="post" enctype="multipart/form-data">
	<input type="file" name="file1" /><br />
    <input type="file" name="file2" /><br />
    <input type="submit" value="提交" />
</form>
<%
	if ($("action")) {
		var jsUpload = new $.io.Upload();
		
		try {
			jsUpload.read();
			jsUpload.saveToFile("file1", Server.MapPath("/"), jsUpload.getAutoFileName());
		} catch (e) {
			Response.Write("上传文件出错" + e.message);
		} finally {
			jsUpload.close();
		}

		Response.Write("上传完成");
	}
%>
</body>
</html>