<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
dim action
action=request.QueryString("action")
select case action
case "del"
if request("shenhe").count=0 then
response.write "<script language=javascript>alert('您没有选择要删除的评论？');history.go(-1);</script>"
response.End
end if
conn.execute ("delete from review where pinglunid in ("&request("shenhe")&")")
response.write "<script language=javascript>alert('批量删除成功!');history.go(-1);</script>"
response.end
case "shenhe"
if request("shenhe").count=0 then
response.write "<script language=javascript>alert('您没有选择要审核的评论？');history.go(-1);</script>"
response.End
end if
conn.execute "update review set shenhe=1 where pinglunid in ("&request("shenhe")&")"
response.write "<script language=javascript>alert('批量审核成功!');history.go(-1);</script>"
response.end
case "delzhou"
dim theday
theday=date-7
conn.execute ("delete from review where pinglundate<#"&theday&"# and shenhe=0")
response.write "<script language=javascript>alert('一周前未审核评论删除成功!');history.go(-1);</script>"
response.end
case "delall"
conn.execute ("delete from review where shenhe=0")
response.write "<script language=javascript>alert('所有未审核评论删除成功!');history.go(-1);</script>"
response.end
end select
%>
