
<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if%>
<%if request.QueryString("action")="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dept",conn,1,3
rs("jsqtoday")=trim(request("jsqtoday"))
rs("jsq")=trim(request("jsq"))
rs.update
rs.close
set rs=nothing
response.write "<script language=javascript>alert('�޸ĳɹ���');history.go(-1);</script>"
response.End
end if
%>
<html>
<head><title><%=webname%>����ͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dept",conn,1,1%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form name="form1" method="post" action="stat.asp?action=save">
<tr> 
<td colspan="2" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�̳Ƿ���������</font></b></td>
</tr>
<tr >
<td align="center">�����̳Ƿ�������</td>
<td style="PADDING-LEFT: 10px"> 
<input name="jsqtoday" type="text" id="458" size="28" value=<%=trim(rs("jsqtoday"))%>>
</td>
</tr>
<tr > 
<td align="center">�̳ǵ��ܷ�������</td>
<td style="PADDING-LEFT: 10px">
<input name="jsq" type="text" id="458url" size="28" value=<%=trim(rs("jsq"))%>>
</td>
</tr>
<tr > 
<td align="center"></td>
<td style="PADDING-LEFT: 10px"> 
<input type="submit" name="Submit" value="�ύ����">
</td>
</tr>
</form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>