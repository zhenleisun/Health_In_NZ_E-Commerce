<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--���ʽ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;">
<!--#include file="include/head.asp"-->
<%
dim action
action=request.QueryString("action")
if InStr(action,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
%>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td ><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" ></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td align="center"><img src="images/body/left06.gif"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href="help.asp?action=gouwuliucheng">��������</a> </td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href="help.asp?action=fukuan">���ʽ</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=tiaokuan>��������</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=yunshushuoming>������֪</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=baomi>���ܰ�ȫ</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=shiyongfalv>��Ȩ����</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=lxwm>��ϵ����</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=about>��������</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=shouhoufuwu>��Ʒ����/�ۺ�</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=gongzuoshijian>�����ϰ�ʱ��</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=baozhuangbiaozhun>���˰�װ��׼</a></td>
                  </tr>
				  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">��<img src="images/body/dian2.gif"> <a href=help.asp?action=baozhuangshuoming>��װ˵��</a></td>
                  </tr>
                  <tr>
                    <td height="24"><img src="images/leftendbg.jpg" width="190" height="36"></td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
    <td width="812" valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" id="AutoNumber3" height="0" width="100%">
      <tr>
        <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="165" valign="top" bgcolor="#FFFFFF"><%select case action
		  case "fukuan"
		  %>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> ���ʽ </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select huikuanfangshi from webinfo",conn,1,1
				response.write trim(rs("huikuanfangshi"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "gouwuliucheng"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> ��������</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gouwuliucheng from webinfo",conn,1,1
				response.write trim(rs("gouwuliucheng"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                  <%case "feiyong"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>�ͻ���ʽ������ </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table1">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select songhuofeiyong from webinfo",conn,1,1
				response.write trim(rs("songhuofeiyong"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "yunshushuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> ������֪</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table2">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select yunshushuoming from webinfo",conn,1,1
				response.write trim(rs("yunshushuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "tiaokuan"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> ��������</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table3">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select jiaoyitiaokuan from webinfo",conn,1,1
				response.write trim(rs("jiaoyitiaokuan"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "yunshushuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> ����˵��</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table4">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select yunshushuoming from webinfo",conn,1,1
				response.write trim(rs("yunshushuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "gongzuoshijian"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >><a href=help.asp?action=gongzuoshijian>�����ϰ�ʱ��</a></td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table5">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gongzuoshijian from webinfo",conn,1,1
				response.write trim(rs("gongzuoshijian"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "about"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >><a href=help.asp?action=about>��������</a></td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table5">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select about from webinfo",conn,1,1
				response.write trim(rs("about"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "shouhoufuwu"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>��Ʒ����/�ۺ� </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table6">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select shouhoufuwu from webinfo",conn,1,1
				response.write trim(rs("shouhoufuwu"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "baomi"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>���ܰ�ȫ</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table7">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baomi from webinfo",conn,1,1
				response.write trim(rs("baomi"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "lxwm"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>��ϵ����</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table8">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select lxwm from webinfo",conn,1,1
				response.write trim(rs("lxwm"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "shiyongfalv"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>��Ȩ���� </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select shiyongfalv from webinfo",conn,1,1
				response.write trim(rs("shiyongfalv"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "baozhuangbiaozhun"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>���˰�װ��׼ </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baozhuangbiaozhun from webinfo",conn,1,1
				response.write trim(rs("baozhuangbiaozhun"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "baozhuangshuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>��װ˵�� </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baozhuangshuoming from webinfo",conn,1,1
				response.write trim(rs("baozhuangshuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case else%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>�������� </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gouwuliucheng from webinfo",conn,1,1
				response.write trim(rs("gouwuliucheng"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%end select%>              </td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" height="40"></td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
