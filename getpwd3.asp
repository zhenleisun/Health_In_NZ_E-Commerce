<!--#include file=conn.asp -->
<!--#include file=md5.asp -->
<%
username=request("username")
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from [user] where username='"&username&"'"
rs.open sql,conn,1,1
If rs.eof Then
%>
<script language="javascript">
alert("����û���û��ע���أ��뵽��ҳע��ɣ�")
location.href="index.asp"
</script>
<%
End If
If rs("answer")<>md5(trim(request("answer"))) Then
Response.redirect"getpwderr.asp?content=��"
End If
%>
<html><head><title>ȡ������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("�Բ��������������룡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 6)
	{
	  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("��������̫���˰ɣ�")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
}
//-->
</script>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<form method="POST" name="FrmAddLink" language="javascript" onSubmit="return FrmAddLink_onsubmit()" action="getpwd4.asp?username=<%=username%>&depid=<%=depid%>">      
<table border="0" cellpadding="2" cellspacing="1" width="400" bgcolor="#6FABCB" style="border-collapse: collapse" align="center">
<TR><TD bgColor=#ffffff height=1></TD></TR>
<TR><TD align=center>�������������������롣</TD>
</TR>
<tr> 
<td bgcolor="F8FAFB">
<p align="center">�� �� �룺
<input class=wenbenkuang type="password" name="passwd" size="15" maxlenght=18><br>
��֤���룺
<input class=wenbenkuang type="password" name="passwd2" size="15" maxlenght=18><br>
<input class=go-wenbenkuang type="submit" value=" �������� " name="submit"></p>
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></p>
</td>
</tr>
</table>
</form>
</body>
</html>
<%   
conn.close   
set conn=nothing 
%>