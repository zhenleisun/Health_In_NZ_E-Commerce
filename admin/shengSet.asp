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

dim shengid,shengorder,shengname,shengno,i
if request("action")="update" then
    for i=1 to request.form("shengid").count
	shengid=replace(request.form("shengid")(i),"'","")
	shengorder=replace(request.form("shengorder")(i),"'","")
	shengname=replace(request.form("shengname")(i),"'","")
	shengno=replace(request.form("shengno")(i),"'","")
	if replace(request.form("shengorder")(i),"'","")="" then
%>
<script language=javascript>
history.back()
alert("����дʡ����ʾ����")
</script>
<%
		Response.End
	end if
	conn.execute("update province set shengorder="&shengorder&",shengname='"&shengname&"',shengno='"&shengno&"' where id="&shengid)

    next
conn.close
set conn=nothing
response.write "<script language=javascript>alert('ʡ���óɹ���');window.location.href='shengmanage.asp';</script>"
end if


if request("action")="add" then
	shengno=request.form("shengno")
	shengname=trim(request("shengname"))
	shengorder=request.form("shengorder")
	If shengname="" Then
		response.write "ʡ���Ʋ���Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
	If shengno="" Then
		response.write "ʡ������Ʋ���Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
	If shengorder="" Then
		response.write "������Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from province",conn,1,3
rs.addnew
rs("shengname")=shengname
rs("shengno")=shengno
rs("shengorder")=shengorder
rs.update
rs.close
set rs=nothing

conn.close
set conn=nothing
response.write "<script language=javascript>alert('ʡ���ӳɹ���');window.location.href='shengmanage.asp';</script>"
end if
%>