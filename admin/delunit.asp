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
dim danweiid
danweiid=request("danweiid")
if not isnumeric(danweiid) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
	conn.execute("delete from unit where id=" &danweiid)
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('�ɹ�ɾ����λ��');window.location.href='manageunit.asp';</script>"
	response.end
%>