<!--#include file="conn.asp" -->
<!--#include file="config.asp" -->
<%
username=request.form("username")
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from [user] where username='"&username&"' "
rs.open sql,conn,1,1
If rs.eof Then
%>
<script language="javascript">
alert("����û�û��ע�ᣬ��ע�ᣡ")
;javascrip:close();</script>
<%
End If
%>
<html><head><title>ȡ������ | �ش�����</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function form1_onsubmit() {
if (document.form1.answer.value=="")
	{
	alert("��������������𰸡�")
	document.form1.answer.focus()
	return false
	}
}
// --></script>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<form method="POST" name="form1" language="javascript" onSubmit="return form1_onsubmit()" action="getpwd3.asp?username=<%=username%>">
<table border="0" cellpadding="2" cellspacing="1" width="400" bgcolor="#6FABCB" style="border-collapse: collapse" align="center">
<TR><TD bgColor=#ffffff height=1></TD></TR>
<TR><TD align=center height=30>�ڶ�������ش��������⡣</TD></TR>
<tr bgcolor="#ffffff"> 
<td height="100" bgcolor="F8FAFB">  
<table border="0" cellpadding="0" width="99%" cellspacing="0" height="1" style="border-collapse: collapse" bordercolor="#111111">
<%if rs("quesion")<>"" then%>
<tr> 
<td width="100%" height="22">
�� �⣺
<font color="red"><%=rs("quesion")%></font> 
	<%     
conn.close     
set conn=nothing     
%>
</td>
</tr>
<tr> 
<td width="100%" height="1" align="center">�� ����
<input class=wenbenkuang type="text" name="answer" size="18">
<input class=go-wenbenkuang name="imageField" type="submit" value=" ��һ�� " onFocus="this.blur()">
<br>
���������ʾ������Ϊ�գ�����*<a href="mailto:<%=webemail%>">����Ա</a>*��ϵȡ����������!</td>
</tr>
<%end if%>
</table>
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></p>
</td>
</tr>
</table>
</form>
</div>
</body>
</html>