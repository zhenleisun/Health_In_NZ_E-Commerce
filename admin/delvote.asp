<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
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
response.write "<script language=javascript>alert('ȷ����');window.location.href='managevote.asp';</script>"
%>