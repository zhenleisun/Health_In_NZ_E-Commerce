<%
sub dindan()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<style type="text/css">
<!--
.style1 {	color: #99CC00;
	font-weight: bold;
}
-->
</style>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="1">
	<tr> 
		<td width="100%" align="right"><img src="images/mingle/state.gif" width="20" height="18" />��ѡ����Ҳ�ͬ״̬�µĶ���
		  <select name="zhuangtai" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
		<option value="user.asp?action=dindan&zhuangtai=0" selected>==��ѡ���Ѷ״̬==</option>
		<option value="user.asp?action=dindan&zhuangtai=0" >ȫ������״̬</option>
		<option value="user.asp?action=dindan&zhuangtai=1" >δ���κδ���</option>
		<option value="user.asp?action=dindan&zhuangtai=2" >�û��Ѿ�������</option>
		<option value="user.asp?action=dindan&zhuangtai=3" >�������Ѿ��յ���</option>
		<option value="user.asp?action=dindan&zhuangtai=4" >�������Ѿ�����</option>
		<option value="user.asp?action=dindan&zhuangtai=5" >�û��Ѿ��յ���</option>
		</select>
		</td>
	</tr>
</table>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
	<tr bgcolor="#F1F1F1" align="center"> 
		<td height="12">������</td>
		<td height="12">����</td>
		<td height="12">����</td>
		<td height="12">����״̬</td>
	</tr>
	  <%set rs=server.CreateObject("adodb.recordset")
	   dim zhuangtai
	  zhuangtai=request.QueryString("zhuangtai")
	  if zhuangtai=0 or zhuangtai="" then
	  select case zhuangtai
	  case "0"
	  rs.open "select distinct(dingdan),userzhenshiname,actiondate,shouhuoname,songhuofangshi,zhifufangshi,zhuangtai from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai<6 order by actiondate desc",conn,1,1
	  case ""
	  rs.open "select distinct(dingdan),userzhenshiname,actiondate,shouhuoname,songhuofangshi,zhifufangshi,zhuangtai from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai<5 order by actiondate desc",conn,1,1
	  end select
	  else
	  rs.open "select distinct(dingdan),userzhenshiname,actiondate,shouhuoname,songhuofangshi,zhifufangshi,zhuangtai from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai="&zhuangtai&" order by actiondate",conn,1,1

  end if

  do while not rs.eof
   %>
	<tr bgcolor="#FFFFFF" align="center"> 
		<td><a href="order.asp?dan=<%=trim(rs("dingdan"))%>" target="_blank"><%=trim(rs("dingdan"))%></a></div></td>
		<td> 
		<%dim cnhw,rs2
		'////
		set cnhw=server.CreateObject("adodb.recordset")
		cnhw.open "select sum(zonger) as zonger from orders where dingdan='"&trim(rs("dingdan"))&"' ",conn,1,1
		response.write "<font color=#FF0000>"&cnhw("zonger")&"</font> Ԫ"
		cnhw.close
		set cnhw=nothing%>
		</td>
		<td> 
			<%
		set cnhw=server.CreateObject("adodb.recordset")
		cnhw.open "select feiyong from orders where dingdan='"&trim(rs("dingdan"))&"' ",conn,1,1
		response.write "<font color=#FF0000>"&cnhw("feiyong")&"</font> Ԫ"
		cnhw.close
		set cnhw=nothing%>
		</td>
		<td> 
			<%select case rs("zhuangtai")
		case "1"
		response.write "δ���κδ���"
		case "2"
		response.write "�û��Ѿ�������"
		case "3"
		response.write "�������Ѿ��յ���"
		case "4"
		response.write "�������Ѿ�����"
		case "5"
		response.write "�û��Ѿ��յ���"
		end select%>
		</td>
	</tr>
	  <%
	   rs.movenext
	  loop
	  rs.close
	  set rs=nothing%>
</table><br>
<%
end sub
sub myinfo()
if request.cookies("Cnhww")("username")<>"" then
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
	<tr align="center"> 
	<td colspan="2" align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5%" align="center"></td>
          <td width="95%">�ҵ�ͳ����Ϣ</td>
        </tr>
      </table></td>
	</tr>
   <%
	set cnhw=server.CreateObject("adodb.recordset")
	cnhw.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"' ",conn,1,1
	ky_jifen=cnhw("jifen")
	%>
    <tr bgcolor="#ffffff" align="center"> 
		<td width="21%">�û����ͣ�</td>
		<td width="79%" align="left">
		
		<font color=red> 
			<%if request.Cookies("cnhww")("reglx")=2 then%>
		VIP�û� 
		<%else%>
		��ͨ��Ա 
		<%end if%>
		</font>
			<%if request.Cookies("cnhww")("reglx")=2 then%>
		���ޣ�<%=cnhw("vipdate")%> 
		<%end if%>
		
		</td>
  </tr>
  <tr bgcolor="#ffffff" align="center">  
    <td>��¼������</td>
    <td align="left"><%=cnhw("logins")%></td>
  </tr>
  <tr bgcolor="#ffffff" align="center"> 
    <td>�ۼƻ��֣�</td>
    <td align="left"><%=cnhw("jifen")%></td>
  </tr>
  <%
	set cnhww=server.CreateObject("adodb.recordset")
	cnhww.open "select sum(zonger) as sum_jine from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai<=5",conn,1,1
	%>
  <tr bgcolor="#ffffff" align="center"> 
    <td>�ܹ����</td>
    <td align="left"><%=cnhww("sum_jine")%></td>
  </tr>
  <%cnhww.close
	set cnhww=nothing%>
  <tr bgcolor="#ffffff" align="center"> 
    <td><font color="#FF0000">Ԥ��</font></td>
    <td align="left"><%=cnhw("yucun")%></td>
  </tr>
  <%if request.cookies("Cnhww")("reglx")=2 then%>
  <tr bgcolor="#ffffff" align="center"> 
    <td>VIP���ޣ�</td>
    <td align="left"><%=cnhw("vipdate")%></td>
  </tr>
  <%end if%>
  <%
	set cnhww=server.CreateObject("adodb.recordset")
	cnhww.open "select count(*) as rec_count from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai=6",conn,1,1
	%>
  <tr bgcolor="#ffffff" align="center"> 
    <td>�ղص���Ʒ����</td>
    <td align="left"><%=cnhww("rec_count")%></td>
  </tr>
  <%cnhww.close
	set cnhww=nothing%>
  <%cnhw.close
	set cnhw=nothing%>
</table>
<%else%>
<table width="90%" height="60"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#F1F1F1">
  <tr>
    <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="528" background="images/mingle/shopcart_bg.gif"><img src="images/mingle/shopcart2.gif" width="219" height="50" /></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
      <tr>
        <td height="30"><table width="598"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="105" colspan="3" bgcolor="#FFFFFF"><img src="images/mingle/login_1.gif" width="598" height="105" /></td>
            </tr>
            <tr>
              <td height="120" colspan="3" background="images/mingle/login_bg.gif" bgcolor="#FFFFFF">
			  <%if request.cookies("Cnhww")("username")="" then%>
			  <table width="370" border="0" align="center" cellpadding="0" cellspacing="0">
                 <form name="fkinfo" method="post" action="checkuserlogin.asp">
				  <tr>
                    <td width="10" rowspan="3" background="images/mingle/login_bg3.gif">&nbsp;</td>
                    <td width="75" height="30"><span class="style1">��</span>�ˡ��ţ�</td>
                    <td width="152"><input name="username" class="wenbenkuang" type="text" id="username" maxLength="18" size="18"></td>
                    <td width="133" rowspan="3" align="center"><input type=image  src="images/mingle/login_4.gif" name="imageField" value="������½"  ></td>
                  </tr>
                  <tr>
                    <td height="30"><span class="style1">��</span>�ܡ��룺</td>
                    <td width="152"><input name="userpassword" class="wenbenkuang" type="password" id="userpassword" maxLength="18" size="18">
<input class="wenbenkuang" type="hidden" name="linkaddress" value="<%=request.servervariables("http_referer")%>"></td>
                  </tr>
<tr> 
<td width="75"  height="30"><span class="style1">��</span>��֤�룺</td>
<td width="152">
<input class=wenbenkuang name=verifycode type=text value="<%If GetCode=9999 Then Response.Write "9999"%>" maxLength=4 size=10>
<img src=GetCode.asp></td>
</tr>
              </table>
			  <%else%>
<%response.redirect("index.asp")%>
<%end if%>			  </td>
            </tr>
            <tr>
              <td width="115" height="105" bgcolor="#FFFFFF"><img src="images/mingle/login3.gif" width="115" height="123" /></td>
              <td width="230" height="123" background="images/mingle/login4.gif" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td height="40" colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td height="42">&nbsp;</td>
                  <td><a href="register.asp"><img src="images/mingle/reg.gif" border="0" /></a></td>
                  <td>&nbsp;</td>
                </tr>
              </table></td>
              <td width="253" height="123" background="images/mingle/login5.gif" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td height="40" colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td width="10%" height="42">&nbsp;</td>
                  <td width="82%"><a href="#"  onClick="javascript:window.open('getpwd.asp','shouchang','width=450,height=300');"><img src="images/mingle/find.gif" width="124" height="31" border="0" /></a></td>
                  <td width="8%">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<%end if%>
<br>
<%
end sub
sub jifen()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
	<tr  align="center"> 
    <td align="left" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%" align="center">&nbsp;</td>
        <td width="90%">���ֶһ���Ʒ˵��</td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
      <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>
            <%set rs=server.createobject("adodb.recordset")
                              rs.open "select jfhj from webinfo ",conn,1,1
                              response.write rs("jfhj")
                              rs.close
                              set rs=nothing
                              %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%" align="center">&nbsp;</td>
        <td width="90%">�ҵĻ������</td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
        <table width="95%" height="24" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
        <%
	set cnhw=server.CreateObject("adodb.recordset")
	cnhw.open "select jifen from [user] where username='"&request.cookies("Cnhww")("username")&"' ",conn,1,1
	ky_jifen=cnhw("jifen")
	cnhw.close
	set cnhw=nothing%>
	<td width="50%">�ҵĿ��û��֣�<font color=#FF0000><%=ky_jifen%></font></td>
        <%
	set cnhw=server.CreateObject("adodb.recordset")
	cnhw.open "select sum(zonger) as sum_jine from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai<=5",conn,1,1
	ky_jifen=cnhw("sum_jine")
	cnhw.close
	set cnhw=nothing%>
	<td width="50%">�ۼƹ�����(��������)��<font color=#FF0000><%=ky_jifen%></font> Ԫ</td>
	</tr>
	</table>
    </td>
  </tr>
	  <tr> 
    <td align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="10%" align="center">&nbsp;</td>
          <td width="90%">������Ԥ����</td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
        <table width="100%" height="24" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr><form name="form2" method="post" action="conversion.asp">
	<td>
  ��
      <input name="jifen" type="text" id="jifen" size="10">
���ֻ����Ԥ���
<input type="submit" name="Submit" value="����">
<input name="act" type="hidden" id="act" value="jifen">
    </td></form>
	<form name="form2" method="post" action="conversion.asp">
	<td>
  ��
      <input name="cunkuan" type="text" id="cunkuan" size="10">
  ԪԤ����ɻ���
  <input type="submit" name="Submit" value="����">
  <input name="act" type="hidden" id="act" value="cunkuan">
            </td></form>
	</tr>
	        <tr>
	          <td colspan="2"><div align="center"><font color="#FF0000">����:1Ԫ=2����</font></div></td>
	          <%
	set cnhw=server.CreateObject("adodb.recordset")
	cnhw.open "select sum(zonger) as sum_jine from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai<=5",conn,1,1
	ky_jifen=cnhw("sum_jine")
	cnhw.close
	set cnhw=nothing%>
          </tr>
	</table>
    </td>
  </tr>
</table>
<br>
<%
end sub
sub shoucang()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<script language="JavaScript">
<!--
var newWindow = null
function windowOpener(loadpos)
{	
  newWindow = window.open(loadpos,'newwindow','width=450,height=350,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');
	newWindow.focus();
}
//-->
</script>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
<tr>
  <td width="5%" align=center>&nbsp;</td>
  <td width="95%" align=left>��Ʒ�ղ����� 10 ��</td>
</tr></table>
<%
set rs=server.CreateObject("adodb.recordset")

rs.open "select orders.actionid,orders.bookid,products.bookname,products.shichangjia,products.huiyuanjia,products.dazhe from products inner join  orders on products.bookid=orders.bookid where orders.username='"&request.cookies("Cnhww")("username")&"' and orders.zhuangtai=6",conn,1,1 

%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
  <form action="stowcart.asp" target="newwindow" method=post name=form1 onsubmit="windowOpener('')">
    <tr bgcolor="#FFFFFF" align="center"> 
      <td width=8%>ѡ��</td>
      <td width=42%>��Ʒ����</td>
      <td width=14%>�г���</td>
      <td width=14%>��Ա��</td>
      <td width=8%>ɾ��</td>
    </tr>
    <%do while not rs.eof%>
    <tr bgcolor="#ffffff" align=center> 
      <td><input name=bookid type=checkbox checked value="<%=rs("bookid")%>" ></td>
      <td align=left><a href=products.asp?id=<%=rs("bookid")%> ><%=rs("bookname")%></a></td>
      <td><s><%=rs("shichangjia")%></s>Ԫ</td>
      <td><%=rs("huiyuanjia")%>Ԫ</td>
      <td><a href=stow.asp?action=del&actionid=<%=rs("actionid")%>&ll=1><img src=images/trash.gif width=15 height=17 border=0></a>
      </td>
    </tr>
    <%
rs.movenext
loop
rs.close
set rs=nothing
%>
	<tr bgcolor="#ffffff" align=center>
	<td height=25 colspan=6> 
	<input class="go-wenbenkuang" onFocus="this.blur()" type=submit name="submit" value=" ���빺�ﳵ ">
	</td>
	</tr>
  </form>
</table><br>
<%
end sub
sub savepass()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<script language=JavaScript>
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function passcheck()
{
    if(document.userpass.userpassword.value.length < 6 || document.userpass.userpassword.value.length >20) {
	document.userpass.userpassword.focus();
    alert("���볤�Ȳ��ܲ�����գ���6λ��20λ֮�䣬���������룡");
	return false;
  }
   if(document.userpass.userpassword.value!==document.userpass.userpassword2.value) {
	document.userpass.userpassword.focus();
    alert("�Բ��������������벻һ����");
	return false;
  }
}
</script>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"' ",conn,1,1
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
  <form name="userpass" method="post" action="saveuserinfo.asp?action=savepass">
    <tr> 
      <td colspan=2 align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center">&nbsp;</td>
            <td>�װ��Ŀͻ������Ǳ�֤��������Ϣ�����ϸ��ܣ����Բ��ṩ������������������!</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td width="30%" bgcolor=#ffffff align="right">�� �� ����</td>
      <td width="70%" bgcolor=#ffffff><font color=#FF0000><%=request.cookies("Cnhww")("username")%></font></td>
	</tr>
    <tr> 
      <td bgcolor=#ffffff align="right">�� �� �룺</td>
      <td bgcolor=#ffffff><input name="userpassword" class="wenbenkuang"; type="password" value="" size="18">
	  <font color="#FF0000">**</font> ���޸���Ϊ��</td>
    </tr>
    <tr> 
      <td bgcolor=#ffffff align="right">����ȷ�ϣ�</td>
      <td bgcolor=#ffffff><input name="userpassword2" class="wenbenkuang" type="password" value="" size="18">
	  <font color="#FF0000">**</font></td>
    </tr>
    <tr align="center"> 
      <td height=25 bgcolor=#ffffff colspan="2">
	  <input class="go-wenbenkuang" onclick="return passcheck();" type="submit" name="submit" value=" �ύ���� ">
	  <input class="go-wenbenkuang" onclick="ClearReset()" type=reset name="Clear" value=" ������д ">
      </td>
    </tr>
  </form>
</table><br>
<%rs.close
set rs=nothing
end sub
sub userziliao()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<script language=JavaScript>
<%dim sql,i,j
	set rs_s=server.createobject("adodb.recordset")
	sql="select * from province order by shengorder"
	rs_s.open sql,conn,1,1
%>
	var selects=[];
	selects['xxx']=new Array(new Option('��ѡ����С���','xxx'));
<%
	for i=1 to rs_s.recordcount
%>
	selects['<%=rs_s("ShengNo")%>']=new Array(
<%
	set rs_s1=server.createobject("adodb.recordset")
	sql="select * from city where shengid="&rs_s("id")&" order by shiorder"
	rs_s1.open sql,conn,1,1
	if rs_s1.recordcount>0 then 
		for j=1 to rs_s1.recordcount
		if j=rs_s1.recordcount then 
%>
		new Option('<%=trim(rs_s1("shiname"))%>','<%=trim(rs_s1("shiNo"))%>'));
<%		else
%>
		new Option('<%=trim(rs_s1("shiname"))%>','<%=trim(rs_s1("shiNo"))%>'),
<%
		end if
		rs_s1.movenext
		next
	else 
%>
		new Option('','0'));
<%
	end if
	rs_s1.close
	set rs_s1=nothing
	rs_s.movenext
	next
rs_s.close
set rs_s=nothing
%>
	function chsel(){
		with (document.userinfo){
			if(szSheng.value) {
				szShi.options.length=0;
				for(var i=0;i<selects[szSheng.value].length;i++){
					szShi.add(selects[szSheng.value][i]);
				}
			}
		}
	}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function checkuserinfo()
{
 if(document.userinfo.useremail.value.length!=0)
  {
    if (document.userinfo.useremail.value.charAt(0)=="." ||        
         document.userinfo.useremail.value.charAt(0)=="@"||       
         document.userinfo.useremail.value.indexOf('@', 0) == -1 || 
         document.userinfo.useremail.value.indexOf('.', 0) == -1 || 
         document.userinfo.useremail.value.lastIndexOf("@")==document.userinfo.useremail.value.length-1 || 
         document.userinfo.useremail.value.lastIndexOf(".")==document.userinfo.useremail.value.length-1)
     {
      alert("Email��ַ��ʽ����ȷ��");
      document.userinfo.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email����Ϊ�գ�");
   document.userinfo.useremail.focus();
   return false;
   }
   if(checkspace(document.userinfo.userzhenshiname.value)) {
	document.userinfo.userzhenshiname.focus();
    alert("�Բ�������д������ʵ������");
	return false;
  }
   if(checkspace(document.userinfo.sfz.value)) {
	document.userinfo.sfz.focus();
    alert("�Բ�������д��������֤���룡");
	return false;
  }
  if((document.userinfo.sfz.value.length!=15)&&(document.userinfo.sfz.value.length!=18)) {
	document.userinfo.sfz.focus();
    alert("�Բ�������ȷ��д����֤���룡");
	return false;
  } 
  if(checkspace(document.userinfo.shouhuodizhi.value)) {
	document.userinfo.shouhuodizhi.focus();
    alert("�Բ�������д������ϸ��ַ��");
	return false;
  }
  if(checkspace(document.userinfo.youbian.value)) {
	document.userinfo.youbian.focus();
    alert("�Բ�������д�ʱ࣡");
	return false;
  }
  if(document.userinfo.youbian.value.length!=6) {
	document.userinfo.youbian.focus();
    alert("�Բ�������ȷ��д�ʱ࣡");
	return false;
  } 
    if(checkspace(document.userinfo.usertel.value)) {
	document.userinfo.usertel.focus();
    alert("�Բ���������������ϵ�绰��");
	return false;
  }
}
</script>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"' ",conn,1,1
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F1F1F1">
  <form name="userinfo" method="post" action="saveuserinfo.asp?action=userziliao">
    <tr> 
      <td colspan=2 align="left"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center"></td>
            <td>�װ��Ŀͻ������Ǳ�֤��������Ϣ�����ϸ��ܣ������ṩ������������������!</td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td width="30%" align="right">�� �� ����</td>
      <td width="70%"><font color=#FF0000>
	  <%=request.cookies("Cnhww")("username")%></font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">�����ʼ���</td>
      <td> 
        <input name=useremail class="wenbenkuang" type=text value="<%=trim(rs("useremail"))%>">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">�Ƿ񹫿������ַ��</td>
      <td> 
        <input type="radio" name="ifgongkai" value="1" <%if rs("ifgongkai")=1 then%>checked<%end if%>>
        ������
        <input type="radio" name="ifgongkai" value="0" <%if rs("ifgongkai")=0 then%>checked<%end if%>>
        ������</td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">��ʵ������</td>
      <td> 
        <input name=userzhenshiname class="wenbenkuang" type=text value="<%=trim(rs("userzhenshiname"))%>" size="10">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">�Ա� </td>
      <td> 
        <input type="radio" name="shousex" value="1" <%if rs("sex")=1 then%>checked<%end if%>>
        �С�
        <input type="radio" name="shousex" value="0" <%if rs("sex")=0 then%>checked<%end if%>>
        Ů��
        <input type="radio" name="shousex" value="2" <%if rs("sex")=2 then%>checked<%end if%>>
        ����</td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">���ڳ��У�</td>
      <td> 
        <select size="1" class="wenbenkuang" name="szSheng" onChange=chsel()>
          <option value="xxx" selected>��ѡ��ʡ�ݡ���</option>
          <%dim tmpShengid
tmpShengid=0
set rs_s=server.createobject("adodb.recordset")
sql="select * from province  order by shengorder"
rs_s.open sql,conn,1,1
while not rs_s.eof
     if rs("szSheng")=rs_s("ShengNo") then
          tmpShengid=rs_s("id")
%>
          <option value="<%=rs_s("ShengNo")%>" selected ><%=trim(rs_s("ShengName"))%></option>
          <%
     else
%>
          <option value="<%=rs_s("ShengNo")%>" ><%=trim(rs_s("ShengName"))%></option>
          <%
     end if
    rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
        </select>
        <select size="1" class="wenbenkuang" name="szShi">
          <%
set rs_s=server.createobject("adodb.recordset")
sql="select * from city where shengid="&tmpShengid&" order by shiorder"
rs_s.open sql,conn,1,1
while not rs_s.eof
%>
          <option value="<%=rs_s("ShiNo")%>" <%if rs("szShi")=rs_s("ShiNo") then%>selected<%end if%>><%=trim(rs_s("ShiName"))%></option>
          <%
    rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
        </select>
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">��ϸ��ַ��</td>
      <td> 
        <input name=shouhuodizhi type=text class="wenbenkuang" value="<%=trim(rs("shouhuodizhi"))%>" size="30">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">��ϵ�绰��</td>
      <td> 
        <input name=usertel type=text class="wenbenkuang" value="<%=trim(rs("usertel"))%>" size="12">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">�ʱࣺ </td>
      <td> 
        <input name=youbian type=text class="wenbenkuang" value="<%=trim(rs("youbian"))%>" ONKEYPRESS="event.returnValue=IsDigit();" size="12">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">QQ��</td>
      <td> 
        <input name=QQ type=text class="wenbenkuang" value="<%=trim(rs("oicq"))%>" size="12" maxlength="12">      </td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">������ҳ��</td>
      <td> 
        <input name=homepage type=text class="wenbenkuang" value="<%=trim(rs("homepage"))%>" size="30">      </td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">���ҽ��ܣ�</td>
      <td> 
        <textarea name="content" cols="30" rows="5" class="wenbenkuang"><%=trim(rs("content"))%></textarea>      </td>
    </tr>
    <tr align="center"> 
      <td height=25 bgcolor=#ffffff colspan="2">
	  <input class="go-wenbenkuang" onclick="return checkuserinfo();" type="submit" name="submit" value=" �ύ���� ">
	  <input class="go-wenbenkuang" onclick="ClearReset()" type=reset name="Clear" value=" ������д ">      </td>
    </tr>
  </form>
</table>
<br>
<%rs.close
set rs=nothing
end sub
sub famess()
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
%>
<div align=center><!--��ʹ�õ�����Ȥ��Ѱ汾�����������ޡ��Ƽ�ʹ�ùٷ���ʽ���������޹������ƣ�����������ַ�鿴���°湺��ϵͳ<br />
  <a href="http://www.cnhww.com" target="_blank">www.cnhww.com</a><br />
  <a href="http://www.Shop7z.com" target="_blank">www.Shop7z.com </a>--> 
  <table width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#cccccc">
	<tr align=center>
		<td><strong><font color=#FFFFFF>���Ͷ���Ϣ</font></strong></td>
	</tr>
	<form name="add" method="post" action="mymsg_hand1.asp">
 	<tr>
		<td>������Ա����վ����Ϣ��</td>
	</tr>
	<tr>
        <td><textarea rows="15" name="neirong" cols="100" style="BORDER: darkgray 1px solid; FONT-SIZE: 8pt; FONT-FAMILY: verdana ; overflow:auto;"></textarea></td>
    </tr>
	<tr>
		<td height=50 align=center><input type="submit" value="��������" name="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="������д"><input name="send" type="hidden" value="ok"></td>
	</tr>
	</form></table>
</div> 
  <% end sub 
     sub soumess()

sql="select * from sms where del='0' and (name='ALL' or name='"&request.cookies("cnhww")("username")&"') order by name asc, riqi desc"
set rs=Server.CreateObject("ADODB.recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
	response.write "<table width=90% border=0 align=center><tr><td><marquee width=100% height=25>��ʱû��վ����Ϣ</marquee></td></tr></table>"
	else
do while not rs.eof
neirong=rs("neirong")
riqi=rs("riqi")
isnew=rs("isnew")
fname=rs("fname")
id=rs("id")
%>
<form action=my_msg.asp method=post>
<table width="90%" border="1" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF" style="word-break:break-all">
<tr>
<td onMouseOver="bgColor='#e7e7e7'" onMouseOut="bgColor='#FFFFFF'">
<table border=0 width=100%>
<tr>
            <td width=150 rowspan=2> <a href='user.asp?action=famess'><img alt="�ظ�����Ϣ" src=images/small/m_replyp.gif border=0></a> 
              &nbsp; <span  style="cursor:hand" onclick="{if(confirm('��ʾ����Ϣɾ���󲻿ɻָ���\n\n��ȷʵɾ��������Ϣ�� ')){location.href='my_msg.asp?delid=<%=id%>';}}"><img ALT="ɾ������Ϣ" src=images/small/m_delete.gif border=0></span>&nbsp;&nbsp; 
            </td>
            <td>&nbsp; </td>
          </tr>
<tr><td>�����ˣ�<%=fname%>  &nbsp;  ����ʱ�䣺<%=riqi%></td></tr>
<tr><td colspan=2><hr color="#FFFFFF" WIDTH=100%>
<%=replace(neirong,vbCRLF,"<BR>")%>
</td></tr>
</table>
</td></tr>
</table>
</form>
<%

rs.movenext
if rs.eof then exit do
loop
conn.execute "UPDATE sms SET isnew ='1' WHERE name ='"&request.cookies("cnhww")("username")&"'"
response.write "<table width=90% border=0 align=center><tr><td height=50 valign=top>"

response.write "</td></tr></table>"

end if
rs.close
set rs=nothing
end sub %>