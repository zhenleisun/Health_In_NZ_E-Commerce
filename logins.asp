<table width="87%" 
border="0" align="center" cellpadding="0" cellspacing="0">
  <tbody>
    <%if request.cookies("Cnhww")("username")=""  then%>
  </tbody>
  <form action="checkuserlogin.asp" method="post" name="userlogin" id="userlogin">
    <tr>
      <td height="10" colspan="2" class="unnamed2"><div align="center">
        <div align="center">�˿�����,������Ʒ���ȵ�¼</div>
      </div></td>
    </tr>
    <tr>
      <td width="35%" class="unnamed2"><div align="right">�ˡ��ţ�</div></td>
      <td width="65%"><div align="left">
        <input name="username" class="wenbenkuang" type="text" id="username2" maxlength="18" size="12" />
      </div></td>
    </tr>
    <tr>
      <td class="unnamed2"><div align="right">�ܡ��룺</div></td>
      <td><div align="left">
        <input name="userpassword" class="wenbenkuang" type="password" id="userpassword2" maxlength="18" size="12" />
        <input class="wenbenkuang" type="hidden" name="linkaddress2" value="<%=request.servervariables("http_referer")%>" />
        <br />
      </div></td>
    </tr>
    <tr>
      <td class="unnamed2"><div align="right">��֤�룺</div></td>
      <td><div align="left">
        <input class="wenbenkuang" name="verifycode" type="text" value="<%If GetCode=9999 Then Response.Write "9999"%>" maxlength="4" size="6" />
        <img src="GetCode.asp" /></div></td>
    </tr>
    <tr>
      <td height="10" colspan="2"></td>
    </tr>
    <tr>
      <td height="17" colspan="2">
        <div align="center">
          <input  name="imageField" src="images/login.gif" border="0" type="image" onfocus="this.blur()" />
          <a href="register.asp"><img src="images/reg.gif" hspace="5" vspace="0" border="0" /></a></div></td>
    </tr>
    <tr>
      <td height="10" colspan="2"></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><img src="images/dot03.gif" width="9" height="9" hspace="5" /><a href="#"  onClick="javascript:window.open('getpwd.asp','shouchang','width=450,height=300');"><font color=000000>���붪ʧ/�һ�����</font></a></div></td>
    </tr>
  </form>
  <%else%>
  <tr>
    <td colspan="2"class="unnamed2" ><div align="center">
      <%
set rs=server.createobject("adodb.recordset")
rs.open "select jifen,yucun,reglx,vipdate from [user] where username='"&request.cookies("Cnhww")("username")&"'",conn,1,3
if rs("vipdate")<>"" then 
if rs("vipdate")<date and rs("reglx")=2 then
rs("reglx")=1
rs.update
end if
end if
response.cookies("Cnhww")("yucun")=rs("yucun")
response.cookies("Cnhww")("jifen")=rs("jifen")
response.cookies("Cnhww")("reglx=")=rs("reglx")
rs.close
set rs=nothing
if request.cookies("Cnhww")("reglx")=2 then
response.write ""&request.cookies("Cnhww")("username")&" VIP�������<br>��Ŀǰ�� <font color=red>"&request.cookies("Cnhww")("jifen")&"</font> ����<br>Ԥ��� <font color=red>"&request.cookies("Cnhww")("yucun")&"</font> Ԫ "
else
response.write ""&request.cookies("Cnhww")("username")&" ����<br>��Ŀǰ�� <font color=red>"&request.cookies("Cnhww")("jifen")&"</font> ����<br>Ԥ��� <font color=red>"&request.cookies("Cnhww")("yucun")&"</font> Ԫ "
end if
response.write "<br><a href=user.asp><font color=red>�����Ա����</font></a>"
response.write "<br><a href=logout.asp>ע��/�˳�</a>"
%>
    </div></td>
  </tr>
  <%end if%>
</table>
