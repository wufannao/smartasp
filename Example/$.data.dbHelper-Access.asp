<!--#include file="../Src/SmartASP.Core.asp"-->
<!--#include file="../Src/SmartASP.Data.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Access数据库操作</title>
<style type="text/css">
table { width: 500px; }
</style>
</head>

<body>
<table>
<%
	var dbHelper = new $.data.DbHelper(Server.MapPath("Library.mdb"));
	dbHelper.connect();

	if ($("action")) {
		var title = $("title"), introduction = $("introduction");
		if (!title) {
%>
<script language="javascript" type="text/javascript">//<![CDATA
alert("请填写标题");
//]]></script>
<%
		} else {
			dbHelper.addParam("Title", "nvarchar", title, 255);
			dbHelper.addParam("Introduction", "ntext", introduction, introduction.length);
			dbHelper.execute("INSERT INTO Book(Title, Introduction) VALUES(?, ?)");
			title = "";
			introduction = "";
		}
	}
	
	var json = dbHelper.executeJson("SELECT * FROM Book");

	for (var i = 0; i < json.length; i++) {
%>
	<tr>
		<td><%=json[i]["BookId"]%></td>
		<td><%=json[i]["Title"]%></td>
		<td><%=json[i]["Introduction"]%></td>
	</tr>
<%
	}

	dbHelper.close();
%>
</table>
<form action="?action=create" method="post">
	<table>
		<tr>
			<th><label for="title">书名：</label></th>
			<td><input type="text" id="title" name="title" style="width: 200px;" maxlength="50" value="<%=title%>" /></td>
		</tr>
		<tr>
			<th><label for="introduction">介绍：</label></th>
			<td><textarea id="introduction" name="introduction" style="width: 350px;"><%=introduction%></textarea></td>
		</tr>
		<tr>
			<td colspan="2""><input type="submit" value="提 交"></td>
		</tr>
	</table>
</form>

</body>
</html>