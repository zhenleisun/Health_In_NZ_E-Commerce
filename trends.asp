<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="��Ȥ���Ϲ���ϵͳ,��Ȥ���Ϲ���ϵͳʱ�а�,��Ȥ����ϵͳ,���Ϲ���ϵͳ,����ϵͳ,��Ȥ����,�̳�Դ��,�����̵�,�����̵�ϵͳ,����ע��,��������,��ΰ����">
<meta name="keywords" content="��Ȥ���Ϲ���ϵͳ,��Ȥ���Ϲ���ϵͳʱ�а�,��Ȥ����ϵͳ,���Ϲ���ϵͳ,����ϵͳ,��Ȥ����,�̳�Դ��,�����̵�,�����̵�ϵͳ,����ע��,��������,��ΰ����">

<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<!--#include file="include/head.asp"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="39" valign="top"><table width="100%" height="184" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="right" valign="top"><img src="images/body/jiao.gif" width="15" height="17"></td>
      </tr>
      <tr>
        <td height="90"><div align="right"><img src="images/ttts.gif" width="26" height="110" ></div></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table></td>
    <td width="190" valign="top" style="BORDER-bottom:#FF67A0 1px solid;BORDER-left:#FF67A0 1px solid; BORDER-right:#FF67A0 1px solid;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
        <td><!--#include file="searc.asp"--><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td></td>
      </tr>
    </table></td>
    <td width="700" valign="top"><table cellspacing=0 cellpadding=0 width=100% align=center border=0>
      <tbody>
        <tr>
          <td class=b valign=top align=left width=100% ><%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""�Ƿ�����!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""�Ƿ�����!"");location.href=""index.asp"";</script>"
end if%>
              <%dim newsid
newsid=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news where newsid="&newsid,conn,1,3
rs("viewcount")=rs("viewcount")+1
rs.update
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr>
                  <td width="100%" align="center" valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="left" height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> �̳�����</td>
                      </tr>
                      <tr bgcolor="#ffffff">
                        <td  valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td colspan="3" align="center"><br>
                                <font color=#F46404><strong><%=trim(rs("newsname"))%></strong></font><br>
                                <br>
                                �����ˣ�<%=trim(rs("addname"))%> ��� <font color=red><%=rs("viewcount")%></font> 
                                �Ρ��ֺ� <A     onclick="Zoom.style.fontSize='19px';" href="#">��</A> 
                                <A 
            onclick="Zoom.style.fontSize='16px';" 
            href="#">��</A> <A    onclick="Zoom.style.fontSize='12px';" 
            href="#">С</A>�� ����ʱ�䣺<%=year(rs("adddate"))&"��"&month(rs("adddate"))&"��"&day(rs("adddate"))&"��"%> 
                                <a class="b12" href="#"
            onclick="if (window.print) window.print();return false">��ӡ��ҳ</a> 
                                <hr width="500" size="1" noshade>
                                <font color=#F46404><b> </b></font></td>
                            </tr>
                            <tr>
                              <td colspan="3"><br>&nbsp;&nbsp;&nbsp;&nbsp; <DIV id=Content><FONT id=Zoom><%=trim(rs("newscontent"))%></div><br></td>
                            </tr>
                            <tr>
                              <td colspan="3"></td>
                            </tr>
                            <tr>
                              <td colspan="3" align="center" ><br>
                              <table width="85%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
                                  </tr>
                                <tr>
                                  <td height="40">�����ˣ�<%=trim(rs("addname"))%> </td>
                                  <td>����ʱ�䣺<%=year(rs("adddate"))&"��"&month(rs("adddate"))&"��"&day(rs("adddate"))&"��"%></td>
                                  <td height="25">�ѱ���� <font color=red><%=rs("viewcount")%></font> ��</td>
                                </tr>
                              </table></td></tr>
                            <tr>
                              <td colspan="3" align="right" ><a href='javascript:onclick=history.go(-1)'><img src="images/mingle/infbk.gif" width="108" height="24" border="0" ></a>&nbsp;&nbsp;&nbsp;</td>
                            </tr>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            <%rs.close
set rs=nothing%></td>
          <td ></td>
        </tr>
      </tbody>
    </table></td>
    <td width="68"style="BORDER-left:#FF67A0 1px solid;">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
