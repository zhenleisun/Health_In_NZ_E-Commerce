<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%dim username,userpassword,comeurl,verifycode
username=replace(trim(request("username")),"'","")
userpassword=md5(replace(trim(request("userpassword")),"'",""))
verifycode=replace(trim(request("verifycode")),"'","")
if username="" or userpassword="" then
response.write "<script LANGUAGE='javascript'>alert('您的用户名或密码有误！');history.go(-1);</script>"
response.end
end if
if cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
response.Write "<script LANGUAGE='javascript'>alert('请输入正确的验证码！');history.go(-1);</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [user] where username='"&username&"' and userpassword='"&userpassword&"' " ,conn,1,3
if not(rs.bof and rs.eof) then
if userpassword=rs("userpassword") then
response.cookies("Cnhww")("username")=trim(request("username"))
response.cookies("Cnhww")("reglx")=rs("reglx")
response.cookies("Cnhww")("jifen")=rs("jifen")
response.cookies("Cnhww")("jiaoyijine")=rs("jiaoyijine")
rs("lastlogin")=now()
rs("logins")=rs("logins")+1
rs("userlastip")=Request.ServerVariables("REMOTE_ADDR")
rs.Update
rs.Close
set rs=nothing
username=trim(request("username"))
conn.execute("delete from orders where username='"&username&"' and zhuangtai=7")
conn.execute("delete from ordersaward where username='"&username&"' and zhuangtai=7")
if request("linkaddress")="" then
response.redirect request.servervariables("http_referer")
else
response.redirect request("linkaddress")
end if
else
response.write "<script LANGUAGE='javascript'>alert('对不起，您的用户名或密码有误！');history.go(-1);</script>"
end if
else
response.write "<script LANGUAGE='javascript'>alert('对不起！您的用户名或密码有误！');history.go(-1);</script>"
end if
sub loginok()
response.Write "欢迎您 <font color=red>"&request.cookies("Cnhww")("username")&"</font>"
response.redirect "index.asp"
end sub
%>
