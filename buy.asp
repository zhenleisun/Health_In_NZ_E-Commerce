<!--#include file="conn.asp"-->
<%

action=request.QueryString("action")
if request.Cookies("cnhww")("username")<>"" then
username=trim(request.Cookies("cnhww")("username"))
else
if request.Cookies("cnhww")("dingdanusername")="" then
username=now()
username=replace(trim(username),"-","")
username=replace(username,":","")
username=replace(username," ","")
response.Cookies("cnhww")("dingdanusername")=username
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] ",conn,1,3
rs.addnew
rs("username")=username
rs("niming")=1
rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
rs("adddate")=now()
rs("lastlogin")=now()
rs("logins")=1
rs.update
rs.close
set rs=nothing

else
username=request.Cookies("cnhww")("dingdanusername")
end if
end if
bookid=request.QueryString("id")







if InStr(action,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");window.close();</script>"
response.end
end if
if bookid<>"" then
if not isnumeric(bookid) then 
response.write"<script>alert(""�Ƿ�����!"");window.close();</script>"
response.end
else
if not isinteger(bookid) then
response.write"<script>alert(""�Ƿ�����!"");window.close();</script>"
end if
end if
end if
select case action
case "del"
conn.execute "delete from orders where actionid="&request.QueryString("actionid")
if request.QueryString("ll")=22 then
response.redirect "myuser.asp?action=shoucang"
else
response.redirect "buy.asp?action=show"
end if
response.End
case "add"
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from products where bookid="&bookid,conn,1,1
if request.Cookies("cnhww")("reglx")=2 then 
	danjia=rs_s("vipjia")
else
	danjia=rs_s("huiyuanjia")
end if
kucun=rs_s("kucun")
bookname=rs_s("bookname")
shjiaid=rs_s("shjiaid")
rs_s.close
set rs_s=nothing
if kucun<=0 then
response.write "<script language=javascript>alert('��ѡ������Ʒ��"&bookname&"����ʱȱ�����ܷŵ����ﳵ���ѡ��������Ʒ��');location.href='javascript:onclick=history.go(-1)'</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username,bookcount,zonger from orders where username='"&username&"' and bookid="&bookid&" and zhuangtai=7",conn,1,3
if rs.recordcount=1 then
if kucun<(rs("bookcount")+1) then
response.write "<script language=javascript>alert('��ѡ������Ʒ��"&bookname&"����ʱȱ�����ܷŵ����ﳵ���ѡ��������Ʒ��');location.href='javascript:onclick=history.go(-1)'</script>"
response.end
end if
rs("zonger")=(rs("bookcount")+1)*danjia
rs("bookcount")=rs("bookcount")+1
rs.update
rs.close
set rs=nothing
response.Redirect "buy.asp?action=show"
else
rs.close
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username,shjiaid,zhuangtai,zonger,bookcount,niming,chima,yanse from orders",conn,1,3
rs.addnew
rs("bookid")=bookid
rs("username")=username
rs("zhuangtai")=7
rs("bookcount")=1
rs("chima")=request.form("chima")
rs("yanse")=request.form("yanse")
rs("shjiaid")=shjiaid
rs("zonger")=danjia
if request.cookies("Cnhww")("username")="" then
rs("niming")=1
end if
rs.update
rs.close
set rs=nothing
response.Redirect "buy.asp?action=show"
end if
case "show"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--�ҵĹ��ﳵ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="userinfo.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td><!--#include file="shopcart.asp"--></td>
      </tr>
      <tr>
        <td><!--#include file="include/selltop.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width="812" valign="top" bgcolor="#FFFFFF"><TABLE cellSpacing=0 cellPadding=0 width=100% align=center border=0>
      <TBODY>
        <TR><%
		set rs=server.CreateObject("adodb.recordset")
rs.open "select orders.actionid,orders.bookid,orders.bookcount,orders.zonger,orders.shjiaid,products.bookname,products.shichangjia,products.huiyuanjia,products.vipjia,orders.chima,orders.yanse from products inner join  orders on products.bookid=orders.bookid where orders.username='"&username&"' and orders.zhuangtai=7",conn,1,1 
shuliang=rs.recordcount
%>
            <TD class=b vAlign=top> 
              <%
response.write "<table width=100% align=center border=0 cellspacing=0 cellpadding=0  background=images/body/pdbg01.gif><tr><td  height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp>"&webname&"</a> >> �ҵĹ��ﳵ</td</tr></table>"
 
%>
              <% if shuliang=0 then %><br><div align=center>���Ĺ��ﳵ�ǿյģ��빺��лл!</div><%end if %></div></div></div>
            <br>
            <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#f1f1f1" >
                <form name='form1' method='post' action=chcount.asp>
                  <tr bgcolor=#f1f1f1 align=center> 
                    <td width=25%  >��Ʒ����</td>
                    <td width=15% >���� 
                      <%if request.Cookies("cnhww")("reglx")=2 then %>
                      (VIP) 
                      <%else%>
                      (��Ա) 
                      <%end if%>                    </td>
                    <td width=7% >����</td>
                    <td>����</td>
                    <td>��ɫ</td>
                    <td width=15% >�ܼ�</td>
                    <td width=10% >ɾ��</td>
                  </tr>
                  <%
jianshu=0
zongji=0
do while not rs.eof%>
                  <tr bgcolor="#FFFFFF"> 
                    <td width="25%" height="30" align="center" ><a href=products.asp?id=<%=rs("bookid")%> ><%=rs("bookname")%></a> 
                      <input name=bookid type=hidden value="<%=rs("bookid")%>"> 
                      <input name=actionid type=hidden value="<%=rs("actionid")%>"></td>
                    <td width="15%" height="30" align="center" > 
                      <%
	if request.Cookies("cnhww")("reglx")=2 then 
	response.write rs("vipjia")
	else
	response.write rs("huiyuanjia")
	end if%>
                      Ԫ</td>
                    <td width="0%" height="30" align="center" ><input class=wenbenkuang name=bookcount value="<%=rs("bookcount")%>" size=5>                    </td>
                    <td height="30" align="center" ><%=rs("chima")%>&nbsp;</td>
                    <td height="30" align="center" ><%=rs("yanse")%></td>
                    <td align="center" width="15%" height="30" ><font color=red><%=rs("zonger")%></font> 
                      Ԫ</td>
                    <td align="center" width="10%" height="30" ><a href=buy.asp?action=del&actionid=<%=rs("actionid")%>><img src=images/trash.gif border=0></a>                    </td>
                  </tr>
                  <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
rs.movenext
loop
rs.close
set rs=nothing%>
                  <tr bgcolor="#FFFFFF"> 
                    <td height=30 colspan=7 align=center > ���ﳵ������Ʒ��<font color=red><%=shuliang%></font> 
                      �֡�������<font color=red><%=jianshu%></font> �������ƣ�<font color=red><%=zongji%></font> 
                      Ԫ������Ԥ��<%=request.cookies("Cnhww")("yucun")%> Ԫ </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td align="center" height=50 colspan=7><input class="go-wenbenkuang" type="button" name="imageField12" value="��������"  border="0" onFocus="this.blur()" onClick="this.form.action='index.asp';this.form.submit()"> 
                      <input class="go-wenbenkuang" type="button" name="imageField2" value="�޸�����"  border="0" onFocus="this.blur()" onClick="this.form.action='chcount.asp';this.form.submit()"> 
                      <input name="imageField22" class="go-wenbenkuang" type="button" value="��չ��ﳵ" onFocus="this.blur()" onClick="this.form.action='clearcart.asp';this.form.submit()"> 
                      <input name="imageField222" class="go-wenbenkuang" type="button" value="ȥ����̨" onFocus="this.blur()" onClick="this.form.action='mycart.asp';this.form.submit()"></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td height="60" colspan="7" STYLE="PADDING-LEFT: 20px"> �� 
                      ����������������ѡ��������<br>
                      �� �������������ڹ��ﳵ�ڵĲ�Ʒ�������޸ģ�Ȼ���ѡ�޸�����<br>
                      �� �������ȫ��ȡ���Ѷ����ڹ��ﳵ�еĲ�Ʒ�����ѡ��չ��ﳵ<br>
                      �� �����������������Ĳ�Ʒ�����ѡȥ����̨(��Ա���ȵ�½���ǻ�Ա�������ע���Ϊ��Ա)<br> </td>
                  </tr>
                </form>
              </table></TD>
        </TR>
      </TBODY>
    </TABLE></td>
 </tr>
</table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
<%end select%>
