<!--#include file="conn.asp"-->
<%
if request.Cookies("cnhww")("username")<>"" then
username=trim(request.Cookies("cnhww")("username"))
else
username=request.Cookies("cnhww")("dingdanusername")
end if%>
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--我要下订单</title>
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
response.write "<script language=javascript>alert('对不起，您购物车没有商品，请在购物后，再去“结算中心”！');location.href=""index.asp"";</script>"
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
                    <a href=index.asp><%=webname%></a> >> 下定单</td>
                </tr>
              </table>
              <br> <table width="90%" align="center" border="0" cellspacing="1" cellpadding="2"  bgcolor="#F1F1F1">
                <form name='form1' method='post' action="mycart.asp?action=shop1">
                  <tr bgcolor="#F1F1F1" align="center"> 
                    <td width=25% ><strong>商品名称</strong></td>
                    <td width=15%><strong>市场价</strong></td>
                    <td width=15%>单价 
                      <%if request.Cookies("cnhww")("reglx")=2 then %>
                      (VIP) 
                      <%else%>
                      (会员) 
                      <%end if%>                    </td>
                    <td width=15%><strong>数量</strong></td>
                    <td width=10%><strong>总价</strong></td>
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
							元</td>
                    <td align="center" ><%=rs("bookcount")%></td>
                    <td align="center" ><%=rs("zonger")%> 元</td>
                  </tr>
                  <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
rs.movenext
loop
rs.close
set rs=nothing%>
                  <tr bgcolor=#ffffff align=center> 
                    <td height=30 colspan=5> 您的购物车里有商品：<%=shuliang%> 件　总数量：<%=jianshu%> 
                      件　共计：<font color=red><%=zongji%></font> 元　您有预存款：<%=request.cookies("Cnhww")("yucun")%> 
                      元 </td>
                  </tr>
                  <tr bgcolor=#ffffff align=center> 
                    <td height=40 colspan=5><input class="go-wenbenkuang" type="button" name="Submit" value="修改购物车" onClick="this.form.action='buy.asp?action=show';this.form.submit()"> 
                      <input class="go-wenbenkuang" type="submit" name="Submit32" value="OK,下一步">                    </td>
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
                    <a href=index.asp><%=webname%></a> >> 填写收货信息</td>
                </tr>
              </table>
              <br> <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                <tr > 
                  <td bgColor="#F1F1F1" colspan="2" align="center"><strong>请正确填写以下收货信息</font></strong></td>
                </tr>
                <form name="shouhuoxx" method="post" action="mycart.asp?action=shop2" onSubmit="ssxx">
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >收货人真实姓名：</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input name=userid type=hidden value="<%=userid%>"> 
                      <input name="userzhenshiname" class="wenbenkuang" type="text" id="userzhenshiname" size="16" value=<%=trim(rs("userzhenshiname"))%>>
                      性别： 
                      <select class="wenbenkuang" name="shousex" id="shousex">
                        <option value=1 <%if rs("sex")="1" then%>selected<%end if%> checked>男</option>
                        <option value=2 <%if rs("sex")="0" then%>selected<%end if%>>女</option>
                        <option value=0 <%if rs("sex")="2" then%>selected<%end if%>>保密</option>
                      </select> </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >详细地址：</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="shouhuodizhi" type="text" id="shouhuodizhi" size="50" value=<%=trim(rs("shouhuodizhi"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >邮政编码：</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="youbian" type="text" id="youbian" size="16" value="<%=rs("youbian")%>" ONKEYPRESS="event.returnValue=IsDigit();"></td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >联系电话：</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="usertel" type="text" id="usertel" size="16" value=<%=trim(rs("usertel"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >电子邮件：</td>
                    <td width="70%" height="30" style="PADDING-LEFT: 20px" ><input class="wenbenkuang" name="useremail" type="text" id="useremail" size="30" value=<%=trim(rs("useremail"))%>>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >送货方式：</td>
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
          '//////////判断送货方式
          set rs3=server.CreateObject("adodb.recordset")
          rs3.Open "select * from deliver where fangshi=0 order by songidorder",conn,1,1
          response.Write "<select name=songhuofangshi size=5 style='width:180px'>"
          do while not rs3.EOF
          if not(shangmen=3 and trim(rs3("songid"))=3) then          '显示不显示“送货上门”
          	response.Write "<option value="&rs3("songid") 
          	if int(rs("songhuofangshi"))=int(rs3("songid")) then 
          	response.Write " selected>"
          	else
          	response.Write ">"
          	end if
			response.Write trim(rs3("subject"))&"("&rs3("jsmoney")&"元)"&"</option>"          
			end if
			rs3.MoveNext
			loop
			response.Write "</select>"
			rs3.Close
			set rs3=nothing
			%>                    </td>
                  </tr>
                  <tr bgcolor="#ffffff"> 
                    <td width="30%" height="30" align="right" >支付方式：</td>
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
					  '//////////判断支付方式
					  set rs3=server.CreateObject("adodb.recordset")
					  rs3.open "select * from deliver where fangshi=1 order by songidorder",conn,1,1
					  response.Write "<select name=zhifufangshi size=5 style='width:180px'>"
					  do while not rs3.eof
					  if not(fukuan=4 and trim(rs3("songid"))=4) then	'显示不显示“货到付款”
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
                    <td height="40" colspan="2" align=center><input class="go-wenbenkuang" type="button" name="Submit22" value="上一步" onClick="javascript:history.go(-1)"> 
                      <input class="go-wenbenkuang" type="submit" name="Submit4" value="OK,下一步" onclick='return ssxx();'>                    </td>
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
    alert("对不起，请填写收货人姓名！");
	return false;
  }
  if(checkspace(document.shouhuoxx.shouhuodizhi.value)) {
	document.shouhuoxx.shouhuodizhi.focus();
    alert("对不起，请填写收货人详细收货地址！");
	return false;
  }
  if(checkspace(document.shouhuoxx.youbian.value)) {
	document.shouhuoxx.youbian.focus();
    alert("对不起，请填写邮编！");
	return false;
  }
  if(document.shouhuoxx.youbian.value.length!=6) {
	document.shouhuoxx.youbian.focus();
    alert("对不起，请正确填写邮编！");
	return false;
  } 
    if(checkspace(document.shouhuoxx.usertel.value)) {
	document.shouhuoxx.usertel.focus();
    alert("对不起，请留下您的电话！");
	return false;
  }
    if(checkspace(document.shouhuoxx.songhuofangshi.value)) {
	document.shouhuoxx.songhuofangshi.focus();
    alert("对不起，请选择送货方式！");
	return false;
  }
    if(checkspace(document.shouhuoxx.zhifufangshi.value)) {
	document.shouhuoxx.zhifufangshi.focus();
    alert("对不起，请选择支付方式！");
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
      alert("Email地址格式不正确！");
      document.shouhuoxx.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
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
                    <a href=index.asp><%=webname%></a> >> 提交定单</td>
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
                              <td width=23%><strong>商品名称</strong></td>
                              <td width=15%><strong>市 场 价</strong></td>
                              <td width=15%>单价 
                                <%if request.Cookies("cnhww")("reglx")=2 then %>
                                (VIP) 
                                <%else%>
                                (会员) 
                                <%end if%>                              </td>
                              <td width=12%><strong>数 量</strong></td>
                              <td width=19%><strong>邮寄费</strong></td>
                              <td width=16%><strong>总 价</strong></td>
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
                              <td><s><%=rs("shichangjia")%></s>元</td>
                              <td>
                                <%
	if request.Cookies("cnhww")("reglx")=2 then 
	response.write rs("vipjia")
	else
	response.write rs("huiyuanjia")
	end if%>
                                元</td>
                              <td><%=rs("bookcount")%></td>
                              <td>
                                <%
						  songd=request("songhuofangshi")
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from deliver where songid="&songd,conn,1,1
feiy=rs_s("jsmoney")
response.write rs_s("jsmoney")
%>                              </td>
                              <td><%=rs("zonger")%> 元</td>
                            </tr>
                            <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
'算每件商品的邮寄费
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
                              <td height=30 colspan=6 align=center>货款总计：<font color=red><%=zongji%></font> 
                                元</td>
                            </tr>
                          </table>
                          <br> </td>
                      </tr>
                      <tr bgcolor="#ffffff"> 
                        <td width="50%" rowspan="2" align="left" valign="top"><table width="90%" border="0" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                            <tr> 
                              <td colspan="2" bgcolor="#F1F1F1" align="center">您的订单信息</td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td width="22%" align="center">姓名：</td>
                              <td width="78%" style="PADDING-LEFT: 20px"><%=request("userzhenshiname")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">邮编：</td>
                              <td style="PADDING-LEFT: 20px"><%=request("youbian")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">地址：</td>
                              <td style="PADDING-LEFT: 20px"><%=request("shouhuodizhi")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">电话：</td>
                              <td style="PADDING-LEFT: 20px"><%=request("usertel")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">邮箱：</td>
                              <td style="PADDING-LEFT: 20px"><%=request("useremail")%></td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">送货：</td>
                              <td style="PADDING-LEFT: 20px">
                                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("songhuofangshi"),conn,1,1
if rs.eof and rs.bof then
response.write "方式已经被删除"
else
response.write rs("subject")
end if
rs.close
set rs=nothing%>                              </td>
                            </tr>
                            <tr bgcolor="ffffff"> 
                              <td align="center">支付：</td>
                              <td style="PADDING-LEFT: 20px">
                                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("zhifufangshi"),conn,1,1
if rs.eof and rs.bof then
response.write "方式已经被删除"
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
                              <td colspan="3" bgcolor="#F1F1F1" align="center">送货费计算</td>
                            </tr>
                            <%
'计算费用
'先取出参数
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
'先看是不是送货
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&trim(request("songhuofangshi")),conn,1,1
if cint(request("songhuofangshi"))=3 then	'先看是不是送货
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
	'不是送货上门的还要看是什么方式送货
	if cint(request("songhuofangshi"))=1 then	'看是不是"普通平邮"
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
	'算商品基价+邮寄费=费用
	feiyong=jijia+fudongjia
	if request("zhifufangshi")=17 then
          '得到预存款
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select yucun from [user] where username='"&username&"'",conn,1,1
          yucunkuan=rs2("yucun")
          rs2.close
          set rs2=nothing
	  if yucunkuan<feiyong+zongji then
	  	response.write "<script language=javascript>alert('您的预存款不足，请更换支付方式！');history.go(-1);</script>"
	end if
	end if
	end if%>
                            <tr> 
                              <td  align="right" bgcolor="ffffff">您的送货费用计：<%=feiy%> 
                                元<br>
                                其中基价：<%=zongji%> 元</td>
                              <td  align="right" bgcolor="ffffff"><img src="images/mingle/charge.gif" width="35" height="32"><a href="help.asp?action=feiyong" target="_blank" ><font color="2782B3">查看送货费用说明</font></a></td>
                            </tr>
                            <tr> 
                              <td colspan="2" align="left" bgcolor="ffffff" style="PADDING-right: 20px"><font color=red>订单总金额： 
                                <%=feiy+zongji%> 元</font></td>
                            </tr><% session("payall")=feiy+zongji %>
                          </table></td>
                      </tr>
                      <form name="shouhuoxx" method="post" action="mycart.asp?action=ok">
                        <tr bgcolor="#ffffff" align="center"> 
                          <td colspan="2"><br> <table width="90%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td><input type="checkbox" name="fapiao" value="1">
                                  是否要发票？ 
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
                                <td height="30"><font color="2782B3">此订单的附加说明(30字内)</font> 
                                  <input class="wenbenkuang" type="text" name="liuyan" size="35" maxlength="30"></td>
                              </tr>
                              <tr> 
                                <td align="center"><br> <input class="go-wenbenkuang" type="button" name="Submit222" value="上一步" onClick="javascript:history.go(-1)"> 
                                  <input class="go-wenbenkuang" type="submit" name="Submit42" value="完　成"></td>
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
'得到价格，减库存
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
response.Write "<center>您不能重复提交！</center>"
response.End
end if
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td  background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >>定单提交成功</td>
                </tr>
              </table>
              <br> <br> <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#F1F1F1">
                <tr bgcolor="#ffffff"> 
                  <td height="30" colspan="2" align="center" bgColor="#F1F1F1"><strong>恭喜订单提交成功！我们会在第一时间进行处理，请记下您的订单号以备查询。</strong></td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单号：<font color=red><%=dingdan%></font></td>
                </tr>
                <%if request.cookies("Cnhww")("username")<>"" then%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单查询：您可通过“<a href="javascript:;" onClick="javascript:window.open('user.asp','','')"><font color="ff0000">我的账户</font></a>”&gt;&gt;“<a href="javascript:;" onClick="javascript:window.open('user.asp?action=dindan','','')"><font color="ff0000">我的订单</font></a>”查询您的订单状态。</td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">购物积分：请在收货后通过“<a href="javascript:;" onClick="javascript:window.open('user.asp','','')"><font color="ff0000">我的账户</font></a>”&gt;&gt;“<a href="javascript:;" onClick="javascript:window.open('user.asp?action=dindan','','')"><font color="ff0000">我的订单</font></a>”及时更改您的订单状态为“完成”，因为每笔订单的积分只有在订单完成后才能累计到您的购物积分中。                  </td>
                </tr>
                <%else%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单查询：您不是我们的会员，所以您不能查询您的订单状态。</td>
                </tr>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">购物积分：因为您不是我们的会员，所以您不能获得积分奖励。                  </td>
                </tr>
                <%
		  end if
		  if request("zhifufangshi")=17 then %>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">您是通过预存款支付的，我们会尽快地给你发货的！</td>
                  <script language=javascript>alert('您是通过预存款支付的，我们会尽快地给你发货的！');</script>
                </tr>
                <%else
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&trim(request("zhifufangshi")),conn,1,1
if trim(rs("subject"))="货到付款" then
rs.close
set rs=nothing
%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 100px">您是选择的“货到付款”，我们会尽快给您送货的！</td>
                  <script language=javascript>alert('您是选择的“货到付款”，请确认您的订单，我们会尽快给您送货的！');</script>
                </tr>
                <%else
rs.close
set rs=nothing%>
                <tr> 
                  <td height="30" colspan="2" bgcolor="ffffff" style="PADDING-LEFT: 30px">请您及时依照您选择的支付方式进行汇款，一周以内没有汇款此订单自动作废，汇款时请注明您的<font color="#FF0000">订单号</font>！</td>
                </tr>
                <%end if%>
                <%end if%>
                <%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from deliver where songid="&request("zhifufangshi"),conn,1,1
if rs.eof and rs.bof then
'response.write "方式已经被删除"
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
				<td width="60%" height="30" align="right" bgcolor="ffffff" ><img src="images/mingle/index.gif" width="17" height="20"><a href="index.asp">返回首页 
				  </a><img src="images/mingle/close.gif" width="26" height="20"><a href="#" onClick=javascript:window.close()>关闭窗口</a><font color="#999999">&nbsp;</font>            </td>
				<td width="35%" align="right" bgcolor="ffffff" ><font color="#999999">　订单提交时间：<%=shijian%></font></td>
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
