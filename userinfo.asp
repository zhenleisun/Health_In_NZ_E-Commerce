<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table width=190 border=0 align=center cellpadding=0 cellspacing=0>
<tr>
<td>
<%if request.cookies("Cnhww")("username")="" then%>
<table width="87%" 
border="0" align="center" cellpadding="0" cellspacing="0">
  <tbody>
    <%if request.cookies("Cnhww")("username")=""  then%>
  </tbody>
  <form action="checkuserlogin.asp" method="post" name="userlogin" id="userlogin">
    <tr>
      <td height="10" colspan="2" align="center" class="unnamed2"><div align="center">�˿�����,������Ʒ���ȵ�¼</div></td>
    </tr>
    <tr>
      <td width="35%" align="right" class="unnamed2"><div align="right">�ˡ��ţ�</div></td>
      <td width="65%" align="left"><div align="left">
        <input name="username" class="wenbenkuang" type="text" id="username2" maxlength="18" size="12" />
      </div></td>
    </tr>
    <tr>
      <td align="right" class="unnamed2"><div align="right">�ܡ��룺</div></td>
      <td align="left"><div align="left">
        <input name="userpassword" class="wenbenkuang" type="password" id="userpassword2" maxlength="18" size="12" />
        <input class="wenbenkuang" type="hidden" name="linkaddress2" value="<%=request.servervariables("http_referer")%>" />
        <br />
      </div></td>
    </tr>
    <tr>
      <td align="right" class="unnamed2"><div align="right">��֤�룺</div></td>
      <td align="left"><div align="left">
        <input class="wenbenkuang" name="verifycode" type="text" value="<%If GetCode=9999 Then Response.Write "9999"%>" maxlength="4" size="6" />
        <img src="GetCode.asp" /></div></td>
    </tr>
    <tr>
      <td height="10" colspan="2"></td>
    </tr>
    <tr>
      <td height="17" colspan="2" align="center"><div align="center">
                <input  name="imageField" src="images/login.gif" border="0" type="image" onfocus="this.blur()" />
            <a href="register.asp"><img src="images/reg.gif" hspace="5" vspace="0" border="0" /></a></div></td>
    </tr>
    <tr>
      <td height="10" colspan="2"></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><div align="center"><img src="images/dot03.gif" width="9" height="9" hspace="5" /><a href="#"  onClick="javascript:window.open('getpwd.asp','shouchang','width=450,height=300');"><font color=#000000 >���붪ʧ/�һ�����</font></a></div></td>
    </tr>
  </form>
  <%else%>
  <tr>
    <td colspan="2"class="unnamed2" ></td>
  </tr>
  <%end if%>
</table>
<%else%>
      <table width="90%" align="center" border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=myinfo">�ҵ���Ϣ</a></td>
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=shoucang">�ҵ��ղ�</a></td>
        </tr>
        <tr> 
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=userziliao">��������</a></td>
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=savepass">�޸�����</a></td>
        </tr>
        <tr> 
          <td align="center"><img src="images/dot03.gif" /></td>
          <td align="left"><a href="user.asp?action=jifen">����ת��</a></td>
          <td align="center"><img src="images/dot03.gif" /></td>
          <td align="left"><a href=# onClick="window.open('getpwd.asp ','basket','menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbars=0,resizable=0,width=410,height=175,top=50,left=100,')">�һ�����</a></td>
        </tr>
        <tr> 
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=dindan">�ҵĶ���</a></td>
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href=logout.asp>�˳���¼</a></td>
        </tr>
        <tr> 
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td align="left"><a href="user.asp?action=famess">���Ͷ���</a></td>
          <td width="15%" align="center"><img src="images/dot03.gif" /></td>
          <td width="35%" align="left"><a href="user.asp?action=soumess">���ն���</a></td>
        </tr>
        <%if request.cookies("Cnhww")("reglx")=2 then%>
        <%else%>
        <%end if%>
      </table>
<%end if%></td>
</tr>
</table>
