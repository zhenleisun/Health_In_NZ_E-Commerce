<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<%
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
dim action
action=request.QueryString("action")
select case action
case "savegongao" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select gonggao from webinfo",conn,1,3
rs("gonggao")=trim(request("gonggao"))
rs.update
set rs=nothing
response.Write "<script language=javascript>alert('公告修改成功！');history.go(-1);</script>"
response.End
end select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">更改首页公告</font></b></td>
</tr>
<tr> 
<td height="100" align="center" > 
<form name="form1" method="post" action="affiche.asp?action=savegongao" id="Form1">
<font color="#FF0000">声明：此信息为首页公告，请及时更新。</font>
<table width="80%" border="0" align="center" cellpadding="3" cellspacing="0" id="Table2">
<tr> 
<td align="center"> 
		<textarea name="gonggao" cols="60" rows="10" id="Textarea1"><%set rs=server.CreateObject("adodb.recordset")
                rs.Open "select gonggao from webinfo",conn,1,1
                response.Write trim(rs("gonggao"))
                rs.Close
                set rs=nothing
                %></textarea>
</td>
</tr>
<tr> 
<td height="20" align="center"> 
<input type="submit" name="Submit" value=" 提交保存 " id="Submit1">
<input type="reset" name="Submit2" value=" 重新填写 " id="Reset1">
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
