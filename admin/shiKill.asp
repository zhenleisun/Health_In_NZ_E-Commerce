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


dim shiid,id
shiid=request("shiid")
if shiid<>"" then
if not isnumeric(shiid) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
id=request("id")
	conn.execute("delete from city where id=" &shiid)
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('��ѡ������Ѿ���ɾ����');window.location.href='shimanage.asp';</script>"
	response.end
%>