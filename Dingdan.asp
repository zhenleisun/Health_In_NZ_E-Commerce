<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--商城新闻</title>
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
                    <a href=index.asp><%=webname%></a> >> 订单中心</td>
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
                <td bgcolor="f1f1f1"><b>商品订单查询</b></td>
              </tr>
              <form name="form" method="post" action="dingdan.asp">
                <tr bgcolor="#FFFFFF"> 
                  <td align="center" style='PADDING-LEFT: 5px'>请输入你的查询号： 
                    <input type="text" class="wenbenkuang" name="dan">
                    <input type="submit" name="Submit" value="订单查询">                  </td>
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
			response.write "<script>alert('无此订单,请重新输入！');javascript:history.go(-1);</script>"
			response.end
			end if
			%>
            <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="10">
                  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                    <tr align="center" bgcolor="f1f1f1"> 
                      <td><strong>商品名称</strong></td>
                      <td><strong>订购数量</strong></td>
                      <td><strong>会员价格</strong></td>
                      <td><strong>金额小计</strong></td>
                  </tr>
                  <%zongji=0
		do while not rs.eof%>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'><a href=products.asp?id=<%=rs("bookid")%> ><%=trim(rs("bookname"))%></a></td>
                    <td align="center"><%=rs("bookcount")%></td>
                    <td align="center"><%=rs("danjia")&"元"%></td>
                    <td align="center"><%=rs("zonger")&"元"%></td>
                  </tr>
                <%zongji=rs("zonger")+zongji
		feiyong=rs("feiyong")
		rs.movenext
		loop
		rs.movefirst%>
         <tr bgcolor="#FFFFFF">
          <td colspan="4"align="right">订单总额：<%=zongji%>元＋费用(<%=feiyong%>元)　　共计：<font color="#ff0000">
		  <%=formatnumber(zongji+feiyong,2)%></font>元 
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
                      <td colspan="2" align="left">　<strong>您此次订单号为：[ <%=dingdan%> ]详细信息如下：</strong></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="15%" align="right">订单状态：</td>
                      <td style="PADDING-LEFT: 12px"> <%zhuang()%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">收货人姓名：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("userzhenshiname"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">收货地址：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("shouhuodizhi"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">支付方式：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">联系电话：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("usertel"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">电子邮件：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("useremail"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">送货方式：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">邮编：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("youbian"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">是否要发票：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%if rs("fapiao")=1 then %>
                        <font color=red>需要发票</font> 
                        <%else%>
                        不需要发票 
                        <%end if%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">您的留言：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("liuyan"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">下单日期：</td>
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
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="2" DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%else%>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="3">
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          订单已完成
          <%end if%>
          <%case "2"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%case "3"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%else%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox1" value="checkbox" >
          订单已完成
          <%end if%>
          <%case "4"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="zhuangtai" value="5" >
          用户已经收到货
          <%case "5"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          用户已经收到货
          <%else%>
          <input name="checkbox3" type="checkbox" DISABLED value="checkbox" checked>
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          订单已完成
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