<!--#include file=conn.asp -->
<!--#include file="md5.asp"-->
<%
username=request("username")
passwd=md5(trim(request.form("passwd")))
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from [user] where username='"&username&"'"
rs.open sql,conn,1,3
If rs.eof Then
%>
<script language="javascript">
alert("这个用户还没有注册呢，请到首页注册吧！")
location.href="javascript:history.back()"
</script>
<%
else
rs("userpassword")=passwd
rs.update
end if
rs.close
set rs=nothing
conn.close   
set conn=nothing
%>
<html><head><title>取回密码</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<table width="320" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6FABCB" style="border-collapse: collapse">
<tr>
<td height="100" bgcolor="F8FAFB"> 
<p align="center">新密码设置成功，请返回首页*<a href="index.asp">重新登录</a>*！</p>
</td>
</tr>
<tr>
<td bgcolor="F8FAFB">
<p align="center"><a href="javascript:window.close()" class="1">关闭窗口</a></p></td>
</tr>
</table>
</body>
</html>
