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
dim id
id=request("id")
if id<>"" then
if not isnumeric(id) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
	conn.execute("delete from zipcode where id=" &id)
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('成功删除邮编！');window.location.href='managezipcode.asp';</script>"
	response.end
%>
