<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--�̳�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
  <LINK href="images/css.css" type=text/css rel=stylesheet></head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="190" valign="top">
       <table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="userinfo.asp"--></td>
          </tr>
        </table>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td> <!--#include file="searc.asp"--></td>
          </tr>
		  <tr>
            <td height="15"><img src="images/leftendbg.jpg" width="190" height="36"></td>
          </tr>
        </table>    </td>
    <td width="812" valign="top"><DIV class=banner_mainM1>
        
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td height=28 align="left"  background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >> ��������</td>
                </tr>
              </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
          <tr>
            <td valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
               <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
                <tr>
                    <td align=left><%  if request("dan")="" then %>
            <table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
              <tr align="center"> 
                <td bgcolor="f1f1f1"><b>��Ʒ������ѯ</b></td>
              </tr>
              <form name="form" method="post" action="dingdan.asp">
                <tr bgcolor="#FFFFFF"> 
                  <td align="center" style='PADDING-LEFT: 5px'>��������Ĳ�ѯ�ţ� 
                    <input type="text" class="wenbenkuang" name="dan">
                    <input type="submit" name="Submit" value="������ѯ">                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'>&nbsp;</td>
                </tr>
              </form>
            </table><%end if %>
            <br>
            <%dim dingdan
			dingdan=request("dan")
			if dingdan<>"" then
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select products.bookid,products.shjiaid,products.bookname,products.shichangjia,products.huiyuanjia,orders.actiondate,orders.shousex,orders.danjia,orders.feiyong,orders.fapiao,orders.userzhenshiname,orders.shouhuoname,orders.dingdan,orders.youbian,orders.liuyan,orders.zhifufangshi,orders.songhuofangshi,orders.zhuangtai,orders.zonger,orders.useremail,orders.usertel,orders.shouhuodizhi,orders.bookcount from products inner join orders on products.bookid=orders.bookid where dingdan='"&dingdan&"' ",conn,1,1
			if rs.eof and rs.bof then
			response.write "<script>alert('�޴˶���,���������룡');javascript:history.go(-1);</script>"
			response.end
			end if
			%>
            <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="10">
                  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                    <tr align="center" bgcolor="f1f1f1"> 
                      <td><strong>��Ʒ����</strong></td>
                      <td><strong>��������</strong></td>
                      <td><strong>��Ա�۸�</strong></td>
                      <td><strong>���С��</strong></td>
                  </tr>
                  <%zongji=0
		do while not rs.eof%>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'><a href=products.asp?id=<%=rs("bookid")%> ><%=trim(rs("bookname"))%></a></td>
                    <td align="center"><%=rs("bookcount")%></td>
                    <td align="center"><%=rs("danjia")&"Ԫ"%></td>
                    <td align="center"><%=rs("zonger")&"Ԫ"%></td>
                  </tr>
                <%zongji=rs("zonger")+zongji
		feiyong=rs("feiyong")
		rs.movenext
		loop
		rs.movefirst%>
         <tr bgcolor="#FFFFFF">
          <td colspan="4"align="right">�����ܶ<%=zongji%>Ԫ������(<%=feiyong%>Ԫ)�������ƣ�<font color="#ff0000">
		  <%=formatnumber(zongji+feiyong,2)%></font>Ԫ 
                      &nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
			<tr>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td>
                  <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#cccccc" align="center">
                    <tr bgcolor="f1f1f1"> 
                      <td colspan="2" align="left">��<strong>���˴ζ�����Ϊ��[ <%=dingdan%> ]��ϸ��Ϣ���£�</strong></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%" align="right">����״̬��</td>
                      <td style="PADDING-LEFT: 12px"> <%zhuang()%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�ջ���������</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("userzhenshiname"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�ջ���ַ��</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("shouhuodizhi"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">֧����ʽ��</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">��ϵ�绰��</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("usertel"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�����ʼ���</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("useremail"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�ͻ���ʽ��</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�ʱࣺ</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("youbian"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�Ƿ�Ҫ��Ʊ��</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%if rs("fapiao")=1 then %>
                        <font color=red>��Ҫ��Ʊ</font> 
                        <%else%>
                        ����Ҫ��Ʊ 
                        <%end if%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�������ԣ�</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("liuyan"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">�µ����ڣ�</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=rs("actiondate")%></td>
                    </tr>
                  </table>                </td>
              </tr>
            </table>
            <br>
          <%end if
			sub zhuang()
			select case rs("zhuangtai")
			case "1"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          δ���κδ���<span style='font-family:Wingdings;'>��</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="2" DISABLED>
          �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          �������Ѿ�����<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          �û��Ѿ��յ���
          <%else%>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="3">
          ������ȷ��<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          ���������
          <%end if%>
          <%case "2"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          δ���κδ���<span style='font-family:Wingdings;'>��</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          �������Ѿ�����<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          �û��Ѿ��յ���
          <%case "3"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          δ���κδ���<span style='font-family:Wingdings;'>��</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          �������Ѿ�����<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          �û��Ѿ��յ���
          <%else%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          ������ȷ��<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox1" value="checkbox" >
          ���������
          <%end if%>
          <%case "4"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          δ���κδ���<span style='font-family:Wingdings;'>��</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          �������Ѿ�����<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="zhuangtai" value="5" >
          �û��Ѿ��յ���
          <%case "5"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          δ���κδ���<span style='font-family:Wingdings;'>��</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          �û��Ѿ�������<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          �������Ѿ��յ���<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          �������Ѿ�����<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          �û��Ѿ��յ���
          <%else%>
          <input name="checkbox3" type="checkbox" DISABLED value="checkbox" checked>
          ������ȷ��<span style='font-family:Wingdings;'>��</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          ���������
          <%end if%>
          <%end select
end sub%>
          <br>&nbsp;</td>
                 </tr>
                </table>              </td>
          </tr>
        </table>
    </DIV></td>
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