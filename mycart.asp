<!--#include file="conn.asp"-->
<%
if request.Cookies("cnhww")("username")<>"" then
username=trim(request.Cookies("cnhww")("username"))
else
username=request.Cookies("cnhww")("dingdanusername")
end if%>
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--��Ҫ�¶���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
          <TR> 
            <TD align=left vAlign=top>
              <%dim bookid,action,i
action=request("action")
set rs=server.CreateObject("adodb.recordset")
rs.open "select count(*) as rec_count from orders where username='"&username&"' and zhuangtai=7",conn,1,1
if rs("rec_count")=0 then
response.write "<script language=javascript>alert('�Բ��������ﳵû����Ʒ�����ڹ������ȥ���������ġ���');location.href=""index.asp"";</script>"
response.End
end if
rs.close
set rs=nothing

select case action
case ""
set rs=server.CreateObject("adodb.recordset")

rs.open "select orders.actionid,orders.bookid,orders.bookcount,orders.zonger,products.bookname,orders.shjiaid,products.shichangjia,products.huiyuanjia,products.vipjia from products inner join  orders on products.bookid=orders.bookid where orders.username='"&username&"' and orders.zhuangtai=7",conn,1,1 

%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td  background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >> �¶���</td>
                </tr>
              </table>
              <br> <table width="90%" align="center" border="0" cellspacing="1" cellpadding="2"  bgcolor="#F1F1F1">
                <form name='form1' method='post' action="mycart.asp?action=shop1">
                  <tr bgcolor="#F1F1F1" align="center"> 
                    <td width=25% ><strong>��Ʒ����</strong></td>
                    <td width=15%><strong>�г���</strong></td>
                    <td width=15%>���� 
                      <%if request.Cookies("cnhww")("reglx")=2 then %>
                      (VIP) 
                      <%else%>
                      (��Ա) 
                      <%end if%>                    </td>
                    <td width=15%><strong>����</strong></td>
                    <td width=10%><strong>�ܼ�</strong></td>
                  </tr>
                  <%shuliang=rs.recordcount
					jianshu=0
					zongji=0
					do while not rs.eof%>
                  <tr bgcolor="#ffffff"> 
                    <td height="30" width="25%" align="center"><%=rs("bookname")%> 
                      <input name=bookid type=hidden value="<%=rs("bookid")%>"> 
                      <input name=actionid type=hidden value="<%=rs("actionid")%>">                    </td>
                    <td align="center" ><%=rs("shichangjia")%></td>
                    <td align="center" ><%
						if request.Cookies("cnhww")("reglx")=2 then 
						response.write rs("vipjia")
						else
						response.write rs("huiyuanjia")
						end if%>
							Ԫ</td>
                    <td align="center" ><%=rs("bookcount")%></td>
                    <td align="center" ><%=rs("zonger")%> Ԫ</td>
                  </tr>
                  <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
rs.movenext
loop
rs.close
set rs=nothing%>
                  <tr bgcolor=#ffffff align=center> 
                    <td height=30 colspan=5> ���Ĺ��ﳵ������Ʒ��<%=shuliang%> ������������<%=jianshu%> 
                      �������ƣ�<font color=red><%=zongji%></font> Ԫ������Ԥ��<%=request.cookies("Cnhww")("yucun")%> 
                      Ԫ </td>
                  </tr>
                  <tr bgcolor=#ffffff align=center> 
                    <td height=40 colspan=5><input class="go-wenbenkuang" type="button" name="Submit" value="�޸Ĺ��ﳵ" onClick="this.form.action='buy.asp?action=show';this.form.submit()"> 
                      <input class="go-wenbenkuang" type="submit" name="Submit32" value="OK,��һ��">                    </td>
                  </tr>
                </form>
              </table>
              <%

case "shop1"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&username&"'",conn,1,1
userid=rs("userid")
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td  background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >> ��д�ջ���Ϣ</td>
                </tr>
              </table>
              <br> <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                <tr > 
                  <td bgColor="#F1F1F1" colspan="2" align="center"><strong>����ȷ��д�����ջ���Ϣ</font></strong></td>
                </tr>
                <form name="shouhuoxx" method="post" action="mycart.asp?action=shop2" onSubmit="ssxx">
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >�ջ�����ʵ������</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input name=userid type=hidden value="<%=userid%>"> 
                      <input name="userzhenshiname" class="wenbenkuang" type="text" id="userzhenshiname" size="16" value=<%=trim(rs("userzhenshiname"))%>>
                      �Ա� 
                      <select class="wenbenkuang" name="shousex" id="shousex">
                        <option value=1 <%if rs("sex")="1" then%>selected<%end if%> checked>��</option>
                        <option value=2 <%if rs("sex")="0" then%>selected<%end if%>>Ů</option>
                        <option value=0 <%if rs("sex")="2" then%>selected<%end if%>>����</option>
                      </select> </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >��ϸ��ַ��</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="shouhuodizhi" type="text" id="shouhuodizhi" size="50" value=<%=trim(rs("shouhuodizhi"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >�������룺</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="youbian" type="text" id="youbian" size="16" value="<%=rs("youbian")%>" ONKEYPRESS="event.returnValue=IsDigit();"></td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >��ϵ�绰��</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="usertel" type="text" id="usertel" size="16" value=<%=trim(rs("usertel"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >�����ʼ���</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="useremail" type="text" id="useremail" size="30" value=<%=trim(rs("useremail"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >�ͻ���ʽ��</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" >
                      <%
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * FROM zipcode where youbian='"&trim(request("youbian"))&"'",conn,1,1
          if rs2.recordcount=1 then
          shangmen=1
          else
          shangmen=0
          end if
          rs2.close
          set rs2=nothing
          '//////////�ж��ͻ���ʽ
          set rs3=server.CreateObject("adodb.recordset")
          rs3.Open "select * from deliver where fangshi=0 order by songidorder",conn,1,1
          response.Write "<select name=songhuofangshi size=5 style='width:180px'>"
          do while not rs3.EOF
          if not(shangmen=3 and trim(rs3("songid"))=3) then          '��ʾ����ʾ���ͻ����š�
          	response.Write "<option value="&rs3("songid") 
          	if int(rs("songhuofangshi"))=int(rs3("songid")) then 
          	response.Write " selected>"
          	else
          	response.Write ">"
          	end if
			response.Write trim(rs3("subject"))&"("&rs3("jsmoney")&"Ԫ)"&"</option>"          
			end if
			rs3.MoveNext
			loop
			response.Write "</select>"
			rs3.Close
			set rs3=nothing
			%>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >֧����ʽ��</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" >
                      <%
					set rs2=server.CreateObject("adodb.recordset")
			
					  rs2.Open "select * from deliver ",conn,1,1
					  if cint(request("songhuofangshi"))=3 then
					  fukuan=1
					  else
					  fukuan=0
					  end if
					  rs2.close
					  set rs2=nothing
					  '//////////�ж�֧����ʽ
					  set rs3=server.CreateObject("adodb.recordset")
					  rs3.open "select * from deliver where fangshi=1 order by songidorder",conn,1,1
					  response.Write "<select name=zhifufangshi size=5 style='width:180px'>"
					  do while not rs3.eof
					  if not(fukuan=4 and trim(rs3("songid"))=4) then	'��ʾ����ʾ���������
						response.Write "<option value="&rs3("songid")
						if int(rs("zhifufangshi"))=int(rs3("songid")) then
						response.Write " selected>"
						else
						response.Write ">"
						end if
						response.Write trim(rs3("subject"))&"</option>"
					  end if
					  rs3.movenext
					  loop
					  response.Write "</select>"
					  rs3.close
					  set rs3=nothing
					  %>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td height="40" colspan="2" align=center><input class="go-wenbenkuang" type="button" name="Submit22" value="��һ��" onClick="javascript:history.go(-1)"> 
                      <input class="go-wenbenkuang" type="submit" name="Submit4" value="OK,��һ��" onclick='return ssxx();'>                    </td>
                  </tr>
                </form>
              </table>
              <SCRIPT LANGUAGE="JavaScript">
<!--
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
function ssxx()
{
   if(checkspace(document.shouhuoxx.userzhenshiname.value)) {
	document.shouhuoxx.userzhenshiname.focus();
    alert("�Բ�������д�ջ���������");
	return false;
  }
  if(checkspace(document.shouhuoxx.shouhuodizhi.value)) {
	document.shouhuoxx.shouhuodizhi.focus();
    alert("�Բ�������д�ջ�����ϸ�ջ���ַ��");
	return false;
  }
  if(checkspace(document.shouhuoxx.youbian.value)) {
	document.shouhuoxx.youbian.focus();
    alert("�Բ�������д�ʱ࣡");
	return false;
  }
  if(document.shouhuoxx.youbian.value.length!=6) {
	document.shouhuoxx.youbian.focus();
    alert("�Բ�������ȷ��д�ʱ࣡");
	return false;
  } 
    if(checkspace(document.shouhuoxx.usertel.value)) {
	document.shouhuoxx.usertel.focus();
    alert("�Բ������������ĵ绰��");
	return false;
  }
    if(checkspace(document.shouhuoxx.songhuofangshi.value)) {
	document.shouhuoxx.songhuofangshi.focus();
    alert("�Բ�����ѡ���ͻ���ʽ��");
	return false;
  }
    if(checkspace(document.shouhuoxx.zhifufangshi.value)) {
	document.shouhuoxx.zhifufangshi.focus();
    alert("�Բ�����ѡ��֧����ʽ��");
	return false;
  }
  if(document.shouhuoxx.useremail.value.length!=0)
  {
    if (document.shouhuoxx.useremail.value.charAt(0)=="." ||        
         document.shouhuoxx.useremail.value.charAt(0)=="@"||       
         document.shouhuoxx.useremail.value.indexOf('@', 0) == -1 || 
         document.shouhuoxx.useremail.value.indexOf('.', 0) == -1 || 
         document.shouhuoxx.useremail.value.lastIndexOf("@")==document.shouhuoxx.useremail.value.length-1 || 
         document.shouhuoxx.useremail.value.lastIndexOf(".")==document.shouhuoxx.useremail.value.length-1)
     {
      alert("Email��ַ��ʽ����ȷ��");
      document.shouhuoxx.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email����Ϊ�գ�");
   document.shouhuoxx.useremail.focus();
   return false;
   }
}
//-->
        </script> 
              <%
rs.close
set rs=nothing
'/////////////////////////////////////////
case "shop2"
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td  background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >> �ύ����</td>
                </tr>
              </table>
              <br> <table width="90%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr bgcolor="#ffffff"> 
                        <td align="center" valign="top" height=50 colspan=2> 
                          <%
		set rs=server.CreateObject("adodb.recordset")
	
rs.open "select orders.actionid,orders.bookid,orders.bookcount,orders.zonger,products.bookname,orders.shjiaid,products.anclassid,products.banci,products.shichangjia,products.huiyuanjia,products.vipjia from products inner join  orders on products.bookid=orders.bookid where orders.username='"&username&"' and orders.zhuangtai=7 order by products.shjiaid",conn,1,1 

%>
                          <br> <table width=100% border=0 align=center cellpadding=2 cellspacing=1 bgcolor="#F1F1F1">
                            <tr align=center bgcolor="#F1F1F1"> 
                              <td width=23%><strong>��Ʒ����</strong></td>
                              <td width=15%><strong>�� �� ��</strong></td>
                              <td width=15%>���� 
                                <%if request.Cookies("cnhww")("reglx")=2 then %>
                                (VIP) 
                                <%else%>
                                (��Ա) 
                                <%end if%>                              </td>
                              <td width=12%><strong>�� ��</strong></td>
                              <td width=19%><strong>�ʼķ�</strong></td>
                              <td width=16%><strong>�� ��</strong></td>
                            </tr>
                            <%shuliang=rs.recordcount
jianshu=0
zongji=0
fudongjia=0
bookname=""
do while not rs.eof
bookname=bookname&"|"&rs("bookname")
%>
                            <tr align="center" bgcolor=#ffffff> 
                              <td height="30" align="left"><div align="center"><%=rs("bookname")%> 
                                  <input name=bookid2 type=hidden value="<%=rs("bookid")%>">
                                  <input name=actionid2 type=hidden value="<%=rs("actionid")%>">
                                </div></td>
                              <td><s><%=rs("shichangjia")%></s>Ԫ</td>
                              <td>
                                <%
	if request.Cookies("cnhww")("reglx")=2 then 
	response.write rs("vipjia")
	else
	response.write rs("huiyuanjia")
	end if%>
                                Ԫ</td>
                              <td><%=rs("bookcount")%></td>
                              <td>
                                <%
						  songd=request("songhuofangshi")
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from deliver where songid="&songd,conn,1,1
feiy=rs_s("jsmoney")
response.write rs_s("jsmoney")
%>                              </td>
                              <td><%=rs("zonger")%> Ԫ</td>
                            </tr>
                            <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
'��ÿ����Ʒ���ʼķ�
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from floater where id="&rs("banci"),conn,1,1 
fudongjia=fudongjia+rs("bookcount")*fudongjia_dj
rs_s.close
set rs_s=nothing
firstshjid=rs("shjiaid")
rs.movenext
if not rs.eof then
nextshjid=rs("shjiaid")
end if
loop
rs.close
set rs=nothing%>
                            <tr bgcolor=#ffffff> 
                              <td height=30 colspan=6 align=center>�����ܼƣ�<font color=red><%=zongji%></font> 
                                Ԫ</td>
                            </tr>
                          </table>
                          <br> </td>
                      </tr>
                      <tr bgcolor="#ffffff"> 
                        <td width="50%" rowspan="2" align="left" valign="top"><table width="90%" border="0" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                            <tr> 
                              <td colspan="2" bgcolor="#F1F1F1" align="center">���Ķ�����Ϣ</td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td width="22%" align="center">������</td>
                              <td width="78%" style="PADDING-LEFT: 20px"><%=request("userzhenshiname")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">�ʱࣺ</td>
                              <td style="PADDING-LEFT: 20px"><%=request("youbian")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">��ַ��</td>
                              <td style="PADDING-LEFT: 20px"><%=request("shouhuodizhi")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">�绰��</td>
                              <td style="PADDING-LEFT: 20px"><%=request("usertel")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">���䣺</td>
                              <td style="PADDING-LEFT: 20px"><%=request("useremail")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">�ͻ���</td>
                              <td style="PADDING-LEFT: 20px">
                                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("songhuofangshi"),conn,1,1
if rs.eof and rs.bof then
response.write "��ʽ�Ѿ���ɾ��"
else
response.write rs("subject")
end if
rs.close
set rs=nothing%>                              </td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">֧����</td>
                              <td style="PADDING-LEFT: 20px">
                                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("zhifufangshi"),conn,1,1
if rs.eof and rs.bof then
response.write "��ʽ�Ѿ���ɾ��"
else
response.write rs("subject")
end if
rs.close
set rs=nothing%>                              </td>
                            </tr>
                          </table></td>
                        <td width="50%" align="center" valign="bottom"><img src="images/mingle/heard.gif" width="260" height="56"></td>
                      </tr>
                      <tr bgcolor="#ffffff"> 
                        <td align="right" valign="bottom"><table width="90%" border="0" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                            <tr> 
                              <td colspan="3" bgcolor="#F1F1F1" align="center">�ͻ��Ѽ���</td>
                            </tr>
                            <%
'�������
'��ȡ������
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from charging ",conn,1,1
t_sm1=rs("sm1")
t_sm2=rs("sm2")
t_sm3=rs("sm3")
t_py1=rs("py1")
t_py2=rs("py2")
t_py3=rs("py3")
t_ems1=rs("ems1")
t_ems2=rs("ems2")
t_ems3=rs("ems3")
'�ȿ��ǲ����ͻ�
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&trim(request("songhuofangshi")),conn,1,1
if cint(request("songhuofangshi"))=3 then	'�ȿ��ǲ����ͻ�
	rs.close
	if request.cookies("Cnhww")("reglx")=2 then
		if zongji>=t_sm3 then 
		jijia=0
		else
		jijia=0
		end if
	else
		if zongji>=t_sm2 then 
		jijia=0
		else
		jijia=0
		end if
	end if
	feiyong=jijia+fudongjia
else
	'�����ͻ����ŵĻ�Ҫ����ʲô��ʽ�ͻ�
	if cint(request("songhuofangshi"))=1 then	'���ǲ���"��ͨƽ��"
		rs.close
		if request.cookies("Cnhww")("reglx")=2 then
			if zongji>=t_py2 then 
			jijia=0
			else
			jijia=0
			end if
			else
			jijia=0
		end if
		else
		rs.close
		if request.cookies("Cnhww")("reglx")=2 then
			if zongji>=t_ems2 then 
			jijia=0
			else
			jijia=0
			end if
			else
			jijia=0
		end if
		end if
	'����Ʒ����+�ʼķ�=����
	feiyong=jijia+fudongjia
	if request("zhifufangshi")=17 then
          '�õ�Ԥ���
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select yucun from [user] where username='"&username&"'",conn,1,1
          yucunkuan=rs2("yucun")
          rs2.close
          set rs2=nothing
	  if yucunkuan<feiyong+zongji then
	  	response.write "<script language=javascript>alert('����Ԥ���㣬�����֧����ʽ��');history.go(-1);</script>"
	end if
	end if
	end if%>
                            <tr> 
                              <td  align="right" bgcolor="ffffff">�����ͻ����üƣ�<%=feiy%> 
                                Ԫ<br>
                                ���л��ۣ�<%=zongji%> Ԫ</td>
                              <td  align="right" bgcolor="ffffff"><img src="images/mingle/charge.gif" width="35" height="32"><a href="help.asp?action=feiyong" target="_blank" ><font color="2782B3">�鿴�ͻ�����˵��</font></a></td>
                            </tr>
                            <tr> 
                              <td colspan="2" align="left" bgcolor="ffffff" style="PADDING-right: 20px"><font color=red>�����ܽ� 
                                <%=feiy+zongji%> Ԫ</font></td>
                            </tr><% session("payall")=feiy+zongji %>
                          </table></td>
                      </tr>
                      <form name="shouhuoxx" method="post" action="mycart.asp?action=ok">
                        <tr bgcolor="#ffffff" align="center"> 
                          <td colspan="2"><br> <table width="90%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td><input type="checkbox" name="fapiao" value="1">
                                  �Ƿ�Ҫ��Ʊ�� 
                                  <input name="userzhenshiname" type="hidden" value=<%=trim(request("userzhenshiname"))%>> 
                                  <input name="shousex" type="hidden" value=<%=trim(request("shousex"))%>> 
                                  <input name="useremail" type="hidden" value=<%=trim(request("useremail"))%>> 
                                  <input name="shouhuodizhi" type="hidden" value=<%=trim(request("shouhuodizhi"))%>> 
                                  <input name="youbian" type="hidden" value=<%=trim(request("youbian"))%>> 
                                  <input name="usertel" type="hidden" value=<%=trim(request("usertel"))%>> 
                                  <input name="songhuofangshi" type="hidden" value=<%=trim(request("songhuofangshi"))%>> 
                                  <input name="zhifufangshi" type="hidden" value=<%=trim(request("zhifufangshi"))%>> 
                                  <input name="feiyong" type="hidden" value=<%=feiy%>> 
                                  <input name="zongji" type="hidden" value=<%=zongji%>> 
                                  <input name=userid type=hidden value="<%=request("userid")%>" > 
                                  <input name="bookname" type="hidden" value=<%=bookname%>></td>
                              </tr>
                              <tr> 
                                <td height="30"><font color="2782B3">�˶����ĸ���˵��(30����)</font> 
                                  <input class="wenbenkuang" type="text" name="liuyan" size="35" maxlength="30"></td>
                              </tr>
                              <tr> 
                                <td align="center"><br> <input class="go-wenbenkuang" type="button" name="Submit222" value="��һ��" onClick="javascript:history.go(-1)"> 
                                  <input class="go-wenbenkuang" type="submit" name="Submit42" value="�ꡡ��"></td>
                              </tr>
                            </table>
                            <br></td>
                        </tr>
                      </form>
                    </table></td>
                </tr>
              </table>
              <%
case "ok"
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&username&"'",conn,1,3
rs("userzhenshiname")=trim(request("userzhenshiname"))
rs("sex")=trim(request("shousex"))
rs("useremail")=trim(request("useremail"))
rs("shouhuodizhi")=trim(request("shouhuodizhi"))
rs("youbian")=trim(request("youbian"))
rs("usertel")=trim(request("usertel"))
rs("songhuofangshi")=trim(request("songhuofangshi"))
rs("zhifufangshi")=trim(request("zhifufangshi"))
rs.update
rs.close
set rs=nothing
if session("xiadan")<>minute(now) then
dim shijian,dingdan
shijian=now()
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from orders where username='"&username&"' and zhuangtai=7",conn,1,3
if request.cookies("Cnhww")("username")<>"" then
dingdan=now()
dingdan=replace(trim(dingdan),"-","")
dingdan=replace(dingdan,":","")
dingdan=replace(dingdan," ","")
else
dingdan=username
end if
do while not rs.eof
'�õ��۸񣬼����
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from products where bookid="&rs("bookid"),conn,1,3
if request.cookies("Cnhww")("reglx")=2 then 
danjia=rs2("vipjia")
else
danjia=rs2("huiyuanjia")
end if
rs2("chengjiaocount")=rs2("chengjiaocount")+rs("bookcount")
rs2("kucun")=rs2("kucun")-rs("bookcount")
rs2.update
rs2.close
set rs2=nothing
rs("actiondate")=shijian
if request("zhifufangshi")=17 then    
rs("zhuangtai")=3
else
rs("zhuangtai")=1
end if
rs("dingdan")=dingdan
rs("youbian")=int(request("youbian"))
rs("shouhuoname")=trim(request("userzhenshiname"))
rs("shouhuodizhi")=trim(request("shouhuodizhi"))
rs("zhifufangshi")=int(request("zhifufangshi"))
rs("songhuofangshi")=int(request("songhuofangshi"))
rs("shousex")=int(request("shousex"))
rs("liuyan")=HTMLEncode2(trim(request("liuyan")))
rs("userzhenshiname")=trim(request("userzhenshiname"))
rs("useremail")=trim(request("useremail"))
rs("usertel")=trim(request("usertel"))
rs("userid")=request("userid")
if request("fapiao")<>1 then 
fapiao=0
else
fapiao=1
end if
rs("fapiao")=fapiao
rs("feiyong")=request("feiyong")
rs("danjia")=danjia
rs.update
rs.movenext
loop
rs.close
set rs=nothing
z_jifen=0
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ordersaward where username='"&username&"' and zhuangtai=7",conn,1,3
do while not rs.eof
rs("actiondate")=shijian
rs("zhuangtai")=5
rs("dingdan")=dingdan
rs("userid")=request("userid")
z_jifen=z_jifen+rs("jifen")
rs.update
rs.movenext
loop
rs.close
set rs=nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&username&"'",conn,1,3
rs("jifen")=rs("jifen")-z_jifen
if request("zhifufangshi")=17 then 
rs("yucun")=rs("yucun")-request("feiyong")-request("zongji")
end if
rs.update
rs.close
set rs=nothing
session("xiadan")=minute(now)
else
response.Write "<center>�������ظ��ύ��</center>"
response.End
end if
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td  background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >>�����ύ�ɹ�</td>
                </tr>
              </table>
              <br> <br> <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                <tr bgcolor="#ffffff"> 
                  <td height="30" colspan="2" align="center" bgColor="#F1F1F1"><strong>��ϲ�����ύ�ɹ������ǻ��ڵ�һʱ����д�������������Ķ������Ա���ѯ��</strong></td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">�����ţ�<font color=red><%=dingdan%></font></td>
                </tr>
                <%if request.cookies("Cnhww")("username")<>"" then%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">������ѯ������ͨ����<a href="javascript:;" onClick="javascript:window.open('user.asp','','')"><font color="ff0000">�ҵ��˻�</font></a>��&gt;&gt;��<a href="javascript:;" onClick="javascript:window.open('user.asp?action=dindan','','')"><font color="ff0000">�ҵĶ���</font></a>����ѯ���Ķ���״̬��</td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">������֣������ջ���ͨ����<a href="javascript:;" onClick="javascript:window.open('user.asp','','')"><font color="ff0000">�ҵ��˻�</font></a>��&gt;&gt;��<a href="javascript:;" onClick="javascript:window.open('user.asp?action=dindan','','')"><font color="ff0000">�ҵĶ���</font></a>����ʱ�������Ķ���״̬Ϊ����ɡ�����Ϊÿ�ʶ����Ļ���ֻ���ڶ�����ɺ�����ۼƵ����Ĺ�������С�                  </td>
                </tr>
                <%else%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">������ѯ�����������ǵĻ�Ա�����������ܲ�ѯ���Ķ���״̬��</td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">������֣���Ϊ���������ǵĻ�Ա�����������ܻ�û��ֽ�����                  </td>
                </tr>
                <%
		  end if
		  if request("zhifufangshi")=17 then %>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">����ͨ��Ԥ���֧���ģ����ǻᾡ��ظ��㷢���ģ�</td>
                  <script language=javascript>alert('����ͨ��Ԥ���֧���ģ����ǻᾡ��ظ��㷢���ģ�');</script>
                </tr>
                <%else
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&trim(request("zhifufangshi")),conn,1,1
if trim(rs("subject"))="��������" then
rs.close
set rs=nothing
%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">����ѡ��ġ�������������ǻᾡ������ͻ��ģ�</td>
                  <script language=javascript>alert('����ѡ��ġ������������ȷ�����Ķ��������ǻᾡ������ͻ��ģ�');</script>
                </tr>
                <%else
rs.close
set rs=nothing%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 30px">������ʱ������ѡ���֧����ʽ���л�һ������û�л��˶����Զ����ϣ����ʱ��ע������<font color="#FF0000">������</font>��</td>
                </tr>
                <%end if%>
                <%end if%>
                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("zhifufangshi"),conn,1,1
if rs.eof and rs.bof then
'response.write "��ʽ�Ѿ���ɾ��"
else
'response.write rs("subject")
dim payway
 
end if
rs.close
set rs=nothing
 
  %>
                <tr> 
                  <td height="30" colspan="2" bgcolor="#FFFFFF" STYLE='PADDING-LEFT: 20px'> 
                    <div align="center" class="style1"><a href="pay/check.asp" target="_blank" class="style3"> 
                      <%
						exec="select * from webinfo"
     					set rs=server.createobject("adodb.recordset")
      					rs.open exec,conn,1,3
      					checkid=rs("paytype")
						 if checkid="3" then
 						elseif checkid="1" then
 %>
                      <img src="images/zhifu_westpay.gif" width="237" height="81" border="0"> 
                      <% elseif  checkid="4" then %>
                      <img src="images/zhifu_99bill.gif" width="237" height="81" border="0"> 
                      <% elseif  checkid="6" then %>
                      <img src="images/zhifu_y.gif" width="237" height="81" border="0"> 
                     
                      <%else%>
                      <img src="images/zhifu_wangyin.gif" width="237" height="81" border="0"> 
                      <%end if %>
                    </a>                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="60%" height="30" align="right" bgcolor="ffffff" ><img src="images/mingle/index.gif" width="17" height="20"><a href="index.asp">������ҳ 
				  </a><img src="images/mingle/close.gif" width="26" height="20"><a href="#" onClick=javascript:window.close()>�رմ���</a><font color="#999999">&nbsp;</font>            </td>
				<td width="35%" align="right" bgcolor="ffffff" ><font color="#999999">�������ύʱ�䣺<%=shijian%></font></td>
				<td width="5%" align="center" bgcolor="ffffff" >&nbsp;</td>
			  </tr>
			</table>
			</td>
          </tr>
      </table>
    <%
	response.cookies("Cnhww")("dingdanusername")=""
	end select%></TD>
  </TR>
      </TBODY>
</TABLE>
</td>
  </tr>
</table>
<table width="1002" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>

<!--#include file="include/foot.asp"-->
</body>
</html>
<script language=javascript>
<!--
function regInput(obj, reg, inputStr)
{
	var docSel	= document.selection.createRange()
	if (docSel.parentElement().tagName != "INPUT")	return false
	oSel = docSel.duplicate()
	oSel.text = ""
	var srcRange	= obj.createTextRange()
	oSel.setEndPoint("StartToStart", srcRange)
	var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
	return reg.test(str)
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
   //-->
</script>