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
dim Checked,sql
Checked=request.form("Checked")
if Checked="" then
response.write "<script>alert('对不起，请您选择投票项！');history.go(-1);</Script>"	
	Response.End 	
end if
set rs=server.createobject("adodb.recordset")
sql="Select isChecked from vote where IsChecked=1"
rs.open sql,conn,1,3
if not rs.eof then
	do while not rs.eof
		rs("IsChecked")=0
	rs.movenext
	loop
end if
rs.close
sql="Select IsChecked from vote where ID="&Checked
rs.open sql,conn,1,3
if not rs.EOF then
	do while not rs.EOF 
		rs("isChecked")=1
	rs.MoveNext 
	loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
response.redirect "managevote.asp"
%>
