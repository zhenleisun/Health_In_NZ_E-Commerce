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
<html>
<head>
<title>市管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form method="post" action="shiset.asp?action=update">
<tr> 
<td class="forumRowHighlight" colspan="6" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">市设置</font></b></td>
</tr>
<tr> 
<td class="forumRowHighlight" colspan="6" align="center">[<a href=shimanage.asp>全部</a>]&nbsp;
<%
Set rs_s = Server.CreateObject("ADODB.Recordset")
rs_s.open "SELECT  * From province  order by shengorder",conn,1,1
while not rs_s.eof
response.write "[<a href=shimanage.asp?shengid="&rs_s("id")&">"&rs_s("shengname")&"</a>]&nbsp;"
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
</td>
</tr>
<tr align=center>
<td class="forumRowHighlight" width=10%>序号</td>
<td class="forumRowHighlight" width=20%>所属省</td>
<td class="forumRowHighlight" width=20%>市名称</td>
<td class="forumRowHighlight" width=20%>市编号</td>
<td class="forumRowHighlight" width=20%>排序<br>(必填项，数字)</td>
<td class="forumRowHighlight" width=10%>删除</td>
</tr>
<%
dim shengid
shengid=request("shengid")
Set rs = Server.CreateObject("ADODB.Recordset")
if shengid<>"" then
  rs.open "SELECT  * From city  where shengid="&shengid&" order by shengid,shiorder",conn,1,1
else
  rs.open "SELECT  * From city  order by shengid,shiorder",conn,1,1
end if
if rs.recordcount=0 then 
%>
<tr align=center>
<td class="forumRowHighlight" colspan="6" align="center" width="100%">还没有添加市</td>
</tr>
<%
	else
    do while not rs.eof
%>
<tr align=center>
<td class="forumRowHighlight" align=center><%=rs("ID")%></td>
<td class="forumRowHighlight" align=center> 
<select name="shengid" size="1">
<%
Set rs_s = Server.CreateObject("ADODB.Recordset")
rs_s.open "SELECT  * From province  order by shengorder",conn,1,1
while not rs_s.eof
%>
<option value="<%=rs_s("id")%>" <%if rs_s("id")=rs("shengid") then%>selected<%end if%>><%=rs_s("shengname")%></option>
<%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
</select>
</td>
<td class="forumRowHighlight" align=center> 
<input type="hidden" name="shiid" value="<%=rs("id")%>">
<input class=text type="text" name="shiname" size="10" value="<%=trim(rs("shiname"))%>">
</td>
<td class="forumRowHighlight" align=center> 
<input class=text type="text" name="shino" size="10" value="<%=trim(rs("shino"))%>" ONKEYPRESS="event.returnValue=IsDigit();" >
</td>
<td class="forumRowHighlight" align=center> 
<input class=text type="text" name="shiorder" size="10" value="<%=rs("shiorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" >
</td>
<td class="forumRowHighlight" align=center> 
<a href="shikill.asp?shiid=<%=rs("id")%>&id=<%=shengid%>">删除</a> 
</td>
</tr>
	<%
    rs.MoveNext
    Loop
%>
<tr align=center>
<td class="forumRowHighlight" colspan="8" height=25 align="center" width="100%"> 
<input type="submit" name="Submit2" value="保存修改" style="font-family: 宋体; font-size: 9pt">&nbsp; 
<input type="reset" value="重新设定" style="font-family: 宋体; font-size: 9pt" name="重置" >
<input type="hidden" name="id" value="<%=shengid%>">
</td>
</tr>
<%
end if
rs.close
set rs=nothing
%>
</form>
</table>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td class="forumRowHighlight" colspan="6" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加市</font></b></td>
</tr>
<form method="post" action="shiset.asp?action=add">
<tr align=center>
<td class="forumRowHighlight" width=10% align=center>添加市</td>
<td class="forumRowHighlight" width=20% align=center> 
<select name="shengid" size="1">
<%
Set rs_s = Server.CreateObject("ADODB.Recordset")
rs_s.open "SELECT  * From province  order by shengorder",conn,1,1
while not rs_s.eof
%>
<option value="<%=rs_s("id")%>"><%=rs_s("shengname")%></option>
<%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
conn.close
set conn=nothing
%>
</select>
</td>
<td class="forumRowHighlight" width=20% align=center> 
<input class=text type="text" name="shiname" size="10">
</td>
<td class="forumRowHighlight" width=20% align=center> 
<input class=text type="text" name="shino" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td class="forumRowHighlight" width=20% align=center> 
<input class=text type="text" name="shiorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td class="forumRowHighlight" width=10% align=center> 
<input type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt">
<input type="hidden" name="id" value="<%=shengid%>">
</td>
</tr>
</form>
</table>
</body>
</html>