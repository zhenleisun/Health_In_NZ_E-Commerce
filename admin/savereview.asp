<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
dim action
action=request.QueryString("action")
select case action
case "del"
if request("shenhe").count=0 then
response.write "<script language=javascript>alert('��û��ѡ��Ҫɾ�������ۣ�');history.go(-1);</script>"
response.End
end if
conn.execute ("delete from review where pinglunid in ("&request("shenhe")&")")
response.write "<script language=javascript>alert('����ɾ���ɹ�!');history.go(-1);</script>"
response.end
case "shenhe"
if request("shenhe").count=0 then
response.write "<script language=javascript>alert('��û��ѡ��Ҫ��˵����ۣ�');history.go(-1);</script>"
response.End
end if
conn.execute "update review set shenhe=1 where pinglunid in ("&request("shenhe")&")"
response.write "<script language=javascript>alert('������˳ɹ�!');history.go(-1);</script>"
response.end
case "delzhou"
dim theday
theday=date-7
conn.execute ("delete from review where pinglundate<#"&theday&"# and shenhe=0")
response.write "<script language=javascript>alert('һ��ǰδ�������ɾ���ɹ�!');history.go(-1);</script>"
response.end
case "delall"
conn.execute ("delete from review where shenhe=0")
response.write "<script language=javascript>alert('����δ�������ɾ���ɹ�!');history.go(-1);</script>"
response.end
end select
%>