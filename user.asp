<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--�ҵ�ר��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
 <tr>
	<td width="190" align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
				 <td height="5" background="images/leftbg.jpg"></td>
			 </tr>
			 <tr>
				<td><img src="images/leftendbg.jpg" width="190" height="36"></td>
			 </tr> 
		</table>
	</td>
	<td width="812" valign="top" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <%xAction=lcase(request("action"))
				select case xAction
				case ""
				%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �ҵ���Ϣ</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "myinfo"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �ҵ���Ϣ</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "userziliao"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> ��������</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "savepass"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �޸�����</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "dindan"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �ҵĶ���</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "shoucang"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �ҵ��ղ�</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case "jifen"
	%>
              <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> �ҵĻ���</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%
	case else
	%>
              <tr>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> ר����ҳ</td>
              </tr>
			  <TR>
                <td height=15></td>
              </tr>
              <%end select%>
            </table>
		<!--#include file="user_inc.asp"-->
		<%xAction=lcase(request("action"))
		select case xAction
		case ""
			call myinfo()
		case "myinfo"
			call myinfo()
		case "userziliao"
			call userziliao()
		case "savepass"
			call savepass()
		case "dindan"
			call dindan()
		case "shoucang"
			call shoucang()
		case "jifen"
			call jifen()
		case "viphd"
			call viphd()
		case "sqvip"
			call sqvip()
		case "famess"
			call famess()
		case "soumess"
			call soumess()
		case else
		%>
        <%end select%>
	</td>
    </tr>
	<tr><td height="40" colspan="2">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>


