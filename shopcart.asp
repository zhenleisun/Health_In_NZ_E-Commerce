<table width="190" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td ><img src="images/body/left05.gif" width="190" height="38"  /></td>
</tr>
<tr>
  <td style="background:url(images/leftbg.jpg) repeat-y;"><table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
    <%
	if request.cookies("Cnhww")("username")<>"" then 
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select count(*) as rec_count from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai=7",conn,1,1
	rec_count=rs("rec_count")
	rs.close
	rs.open "select sum(zonger) as zongji from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai=7",conn,1,1
	%>
    <tr>
      <td height="2" align="cneter"></td>
    </tr>
    <tr>
      <td height="20">�����Ĺ��ﳵ��<font color="red"><%=rec_count%></font>����Ʒ</td>
    </tr>
    <tr>
      <td height="20">���ܽ�<font color="red"><%=rs("zongji")%></font>Ԫ</td>
    </tr>
    <tr>
      <td height="20">��<a href="buy.asp?action=show&amp;lx=1" target="_blank">�鿴���ﳵ/����&gt;&gt;</a></td>
    </tr>
    <tr>
      <td height="20">��<a href="user.asp?action=shoucang">�ҵ��ղ�&gt;&gt;</a></td>
    </tr>
    <%rs.close
		set rs=nothing
		else%>
    <tr>
      <td height="20"><div align="center">����<font color="red">û�е�½</font><br />
        ���ﳵ<font color="red">����ʹ��</font> </div></td>
    </tr>
    <tr>
      <td height="1"></td>
    </tr>
    <%end if%>
  </table></td>
</tr>
</table>
