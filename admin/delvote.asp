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
dim id,sql
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="delete from vote where ID="&id
rs.open sql,conn,1,3
conn.close
set conn=nothing
response.write "<script language=javascript>alert('确定！');window.location.href='managevote.asp';</script>"
%>
