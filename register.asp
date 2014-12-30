<html>
<head>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%dim action
action=request.QueryString("action")%>
<title><%=webname%>--新用户注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
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
</script>
<SCRIPT LANGUAGE="JavaScript">
<%dim sql,i,j
	set rs_s=server.createobject("adodb.recordset")
	sql="select * from province order by shengorder"
	rs_s.open sql,conn,1,1
%>
	var selects=[];
	selects['mr']=new Array(new Option('请选择城市……','mr'));
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
<!--
function checkuserinfo()
{
   if(checkspace(document.userinfo.username.value)) {
	document.userinfo.username.focus();
    alert("对不起，请填写用户名！");
	return false;
  }
    if(checkspace(document.userinfo.userpassword.value) || document.userinfo.userpassword.value.length < 6 || document.userinfo.userpassword.value.length >20) {
	document.userinfo.userpassword.focus();
    alert("密码长度不能不能这空，在6位到20位之间，请重新输入！");
	return false;
  }
    if(document.userinfo.userpassword.value != document.userinfo.userpassword1.value) {
	document.userinfo.userpassword.focus();
	document.userinfo.userpassword.value = '';
	document.userinfo.userpassword1.value = '';
    alert("两次输入的密码不同，请重新输入！");
	return false;
  }
 if(document.userinfo.useremail.value.length!=0)
  {
    if (document.userinfo.useremail.value.charAt(0)=="." ||        
         document.userinfo.useremail.value.charAt(0)=="@"||       
         document.userinfo.useremail.value.indexOf('@', 0) == -1 || 
         document.userinfo.useremail.value.indexOf('.', 0) == -1 || 
         document.userinfo.useremail.value.lastIndexOf("@")==document.userinfo.useremail.value.length-1 || 
         document.userinfo.useremail.value.lastIndexOf(".")==document.userinfo.useremail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.userinfo.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.userinfo.useremail.focus();
   return false;
   }
   if(checkspace(document.userinfo.userzhenshiname.value)) {
	document.userinfo.userzhenshiname.focus();
    alert("对不起，请填写真实姓名！");
	return false;
  }
  if(checkspace(document.userinfo.shouhuodizhi.value)) {
	document.userinfo.shouhuodizhi.focus();
    alert("对不起，请填写收货人详细收货地址！");
	return false;
  }
  if(checkspace(document.userinfo.youbian.value)) {
	document.userinfo.youbian.focus();
    alert("对不起，请填写邮编！");
	return false;
  }
  if(document.userinfo.youbian.value.length!=6) {
	document.userinfo.youbian.focus();
    alert("对不起，请正确填写邮编！");
	return false;
  } 
    if(checkspace(document.userinfo.usertel.value)) {
	document.userinfo.usertel.focus();
    alert("对不起，请填写电话号码！");
	return false;
  } 
        if(checkspace(document.userinfo.songhuofangshi.value)) {
	document.userinfo.songhuofangshi.focus();
    alert("对不起，您还没有选择送货方式！");
	return false;
  }
      if(checkspace(document.userinfo.zhifufangshi.value)) {
	document.userinfo.zhifufangshi.focus();
    alert("对不起，您还没有选择支付方式！");
	return false;
  }
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
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                <td><img src="images/leftendbg.jpg"></td>
              </tr>
    </table></td>
    <td width="812" valign="top" bgcolor="#FFFFFF"><TABLE cellSpacing=0 cellPadding=0 width=100% align=center border=0>
      <TBODY>
	  <TR>
                <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> 新用户注册</td>
            </tr>
        <TR>
          <TD class=b vAlign=top align=left>
              <%select case action
				case ""%>
              <table width="95%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
				<tr>
                  <td height="22" bgcolor="#FFFFFF" bordercolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
                      <tr>
                        <td valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td><table width="80%" align="center" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td height="40" align="center"><font color=red><strong>请用户仔细阅读<%=webname%>注册条约</strong></font></td>
                                  </tr>
                              </table></td>
                            </tr>
                            <tr>
                              <td><table width="80%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr bgcolor="#ffffff">
                                    <td><%call tiaoyue()%>                                    </td>
                                  </tr>
                              </table></td>
                            </tr>
                          </table>
                            <table name=agree border="0" cellpadding="10" cellspacing="0" align=center width="80%">
                              <tr align=center>
                                <td width="50%" align="right"><FORM name=register method=post action=register.asp?action=yes>
                                    <INPUT class="go-wenbenkuang" type=submit value=" 同意注册 " name=Submit>
                                </FORM></td>
                                <td width="50%" align="left"><FORM action=index.asp method=post>
                                    <INPUT name="submit" type=submit class=go-wenbenkuang value=" 我不同意 ">
                                </FORM></td>
                              </tr>
                          </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            <%case "yes"%>
              <table width="95%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr>
                  <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="center"><br><p><font color="#FF3300"><b>请您详细填写您的信息<br>
                          您的个人信息对外保密。(带 <font color=red>**</font> 号为必填项) </b></font></p>
                            <form name=userinfo method=post action=register.asp?action=save>
                              <table width="85%" border="0" cellpadding="2" cellspacing="1" bgcolor="#e1e1e1" align="center">
                                <tr>
                                  <td height=15 colspan=2 bgcolor="#f1f1f1"><font color="#FF3300">* 用户名和密码</font> </td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>用 户 名：</td>
                                  <td width=83% class=pad><input class="wenbenkuang" name="username" type="text" id="username" maxLength="18">
                                      <font color=red>**</font>用户名长度不能小于2支持中文</td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>密　　码：</td>
                                  <td class=pad><input class="wenbenkuang" name="userpassword" type="password" id="userpassword" maxLength="18">
                                      <font color=red>**</font>长度必须大于6个字符</td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>确认密码：</td>
                                  <td class=pad><input class="wenbenkuang" name="userpassword1" type="password" id="userpassword1" maxLength="18">
                                      <font color=red>**</font></td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>电子邮箱：</td>
                                  <td class=pad><input class="wenbenkuang" name="useremail" type="text" id="useremail">
                                      <font color=red>**</font></td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td align=right>密码问题：</td>
                                  <td class=pad><input class="wenbenkuang" name="quesion" type="text" id="quesion">
                                    忘记密码的提示问题</td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td align=right>问题答案：</td>
                                  <td class=pad><input class="wenbenkuang" name="answer" type="text" id="answer">
                                    忘记密码的问题答案</td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td colspan="2" valign="middle" bgcolor="#f1f1f1"><font color="#FF3300">* 用户详细资料</font></td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>您的真实姓名：</td>
                                  <td class=pad><input class="wenbenkuang" name="userzhenshiname" type="text" id="userzhenshiname" size="10">
                                      <font color=red>**</font> 以便发货确认 </td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>姓　　别：</td>
                                  <td class=pad><input type=radio name=babysex id=Select1 value=0 checked>
                                    男
                                    <input type=radio name=babysex id=Select1 value=1>
                                    女 </td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td align="right">所在城市：</td>
                                  <td><select size="1" class="wenbenkuang" name="szSheng" onChange=chsel()>
                                      <option value="xxx" selected>请选择省份……</option>
                                      <%dim tmpShengid
										tmpShengid=0
										set rs_s=server.createobject("adodb.recordset")
										sql="select * from province  order by shengorder"
										rs_s.open sql,conn,1,1
										while not rs_s.eof
										   
										%>
                                      <option value="<%=rs_s("ShengNo")%>" selected ><%=trim(rs_s("ShengName"))%></option>
                                      
										<%
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
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>收货(通讯)地址：</td>
                                  <td class=pad><input class="wenbenkuang" name="shouhuodizhi" type="text" id="shouhuodizhi" size="40" maxlength="30">
                                      <font color=red>**</font> </td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>邮　　编：</td>
                                  <td class=pad><input class="wenbenkuang" name="youbian" type="text" id="youbian" maxlength="6" size="10" onKeyPress="event.returnValue=IsDigit();">
                                      <font color=red>**</font> </td>
                                </tr>
                                <tr bgcolor="#ffffff">
                                  <td align="right">QQ：</td>
                                  <td><input name=oicq type=text class="wenbenkuang" id="oicq" size="12" maxlength="12">                                  </td>
                                </tr>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align=right>联系电话：</td>
                                  <td class=pad><input class="wenbenkuang" name="usertel" maxlength="18" type="text" id="usertel">
                                      <font color=red>**</font> </td>
                                </tr>
                                <%
								'////////////送货方式
								response.Write "<tr bgcolor=#FFFFFF><td width=30% align=right>送货方式：</td><td class=pad><select class=wenbenkuang name=songhuofangshi id=songhuofangshi>"
								set rs2=server.CreateObject("adodb.recordset")
								rs2.open "select * from deliver where fangshi=0 order by songidorder",conn,1,1
								do while not rs2.EOF
								response.Write "<option value="&rs2("songid")&">"&trim(rs2("subject"))&"</option>"
								rs2.MoveNext
								loop
								rs2.Close
								response.Write "</select><font color=red>**</font></td></tr>"
								response.Write "<tr bgcolor=#FFFFFF><td width=30% align=right>支付方式：</td><td class=pad><select class=wenbenkuang name=zhifufangshi id=zhifufangshi>"
								'////////////支付方式
								rs2.Open "select * from deliver where fangshi=1 order by songidorder",conn,1,1
								do while not rs2.EOF
								response.Write "<option value="&rs2("songid")&">"&trim(rs2("subject"))&"</option>"
								rs2.MoveNext
								loop
								rs2.Close
								set rs2=nothing
								response.Write "</select><font color=red>**</font></td></tr>"
								%>
                                <tr bgcolor="#FFFFFF">
                                  <td width=17% align="right"></td>
                                  <td class=pad><input class="go-wenbenkuang" onClick="return checkuserinfo();" type=submit name="submit" value=" 提交信息 ">
                                      <input class="go-wenbenkuang" onClick="ClearReset()" type=reset name="Clear" value=" 重新填写 ">                                  </td>
                                </tr>
                              </table>
                            </form></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            <br>
            <%case "save"%>
              <!--#include file="md5.asp"-->
              <%call saveuser()%>
              <%
				end select%>
							  <%sub tiaoyue()
				set rs=server.CreateObject("adodb.recordset")
				rs.Open "select regtiaoyue from webinfo",conn,1,1
				response.Write trim(rs("regtiaoyue"))
				rs.Close
				set rs=nothing
				end sub
				sub saveuser()
				if session("regtimes")=1 then
				response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0 align=center><tr><td height=300 align=center><font color=red>对不起，您刚注册过用户，请稍后再进行注册!</font></td></tr></table>"
				response.End
				end if
				set rs=server.CreateObject("adodb.recordset")
				rs.open "select * from [user] where UserEmail='"&trim(request("useremail"))&"' or UserName='"&trim(request("username"))&"'",conn,1,1
				if rs.recordcount>0 then
				call usererr()
				rs.close
				else
				rs.close
				set rs=server.CreateObject("adodb.recordset")
				rs.open "select * from [user]",conn,1,3
				rs.addnew
				rs("username")=trim(request("username"))
				rs("userpassword")=md5(trim(request("userpassword")))
				rs("useremail")=trim(request("useremail"))
				rs("quesion")=trim(request("quesion"))
				rs("answer")=md5(trim(request("answer")))
				rs("userzhenshiname")=trim(request("userzhenshiname"))
				rs("shouhuodizhi")=trim(request("shouhuodizhi"))
				rs("youbian")=trim(request("youbian"))
				rs("oicq")=trim(request("oicq"))
				rs("usertel")=trim(request("usertel"))
				rs("songhuofangshi")=trim(request("songhuofangshi"))
				rs("zhifufangshi")=trim(request("zhifufangshi"))
				rs("adddate")=now()
				rs("lastlogin")=now()
				rs("logins")=1
				rs("reglx")=1
				rs("jifen")=0
				rs("jiaoyijine")=0
				rs("sex")=1
				rs("userlastip")=Request.ServerVariables("REMOTE_ADDR")
				
				rs.update
				rs.close
				set rs=nothing
				set rs=conn.execute("select top 1 userid,face from [user] order by userid desc")
				userid=rs(0)
				response.cookies("Cnhww")("username")=trim(request("username"))
				response.cookies("Cnhww")("jiaoyijine")=0
				response.cookies("Cnhww")("jifen")=0
				response.cookies("Cnhww")("reglx")=1
				response.Cookies("shangcheng").expires=date+1
				session("regtimes")=1
				session.Timeout=1
				set rs=server.CreateObject("adodb.recordset")
				rs.Open "select mailaddress,mailusername,mailuserpass,mailname,mailsend from webinfo",conn,1,1
				mailaddress=rs("mailaddress")
				mailusername=rs("mailusername")
				mailuserpass=rs("mailuserpass")
				mailname=rs("mailname")
				mailsend=rs("mailsend")
				rs.close
				set rs=nothing
					on error resume next
					topic="您在" & webname & "的注册资料"
					getpass=trim(request("userpassword"))
					mailbody="<html>"
					mailbody=mailbody & "<title>注册信息</title>"
					mailbody=mailbody & "<body>"
					mailbody=mailbody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
					mailbody=mailbody & "<TD valign=middle align=top>"
					mailbody=mailbody & trim(request("username"))&"，您好：<br><br>"
					mailbody=mailbody & "欢迎您注册" & webname & "网上商城，我们将提供给您最好的服务！<br>"
					mailbody=mailbody & "下面是您在" & webname & "网上商城的注册信息：<br><br>"
					mailbody=mailbody & "注册名："&trim(request("username"))&"<br>"
					mailbody=mailbody & "密  码："&getpass&"<br>"
					mailbody=mailbody & "<br><br>"
					mailbody=mailbody & "<center><font color=red>再次感谢您注册" & webname & "网上商城！</font>"
					mailbody=mailbody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
					mailbody=mailbody & "</body>"
					mailbody=mailbody & "</html>"
				Set JMail=Server.CreateObject("JMail.Message")
					JMail.Charset="gb2312"
					JMail.ContentType = "text/html"
				jmail.from = mailsend
				jmail.silent = true
				jmail.Logging = true
				jmail.FromName = mailname
				jmail.mailserverusername = mailusername
				jmail.mailserverpassword = mailuserpass
				jmail.AddRecipient trim(request("useremail"))
				jmail.body=mailbody
				JMail.Subject=topic
				if not jmail.Send ( mailaddress ) then
				SendMail=""
				else
				SendMail="OK"
				end if
					if SendMail="OK" then
					sendmsg="<p>· 您的注册信息已经发往您的邮箱，请注意查收!</p>"
					else
					sendmsg="<p>· 系统出错，注册信息资料未能发送到您的邮箱!</p>"
					end if
					'response.write mailbody
				'end if
				response.Write "<table width=100% align=center border=0 cellspacing=0 cellpadding=0  bordercolor=#CCCCCC><tr><td bordercolor=#FFFFFF bgcolor=#FFFFFF align=center> "
				response.Write "<table width=450 border=0 align=center cellpadding=0 cellspacing=0><tr><td height=260>"
				response.Write "<p>· <font color=red>用户注册成功!</font></p><p>· 恭喜您注册成为我网站的正式用户，请记好您的用户名及密码!</p>"
				response.Write "<p>· <a href=index.asp>返回首页</a></p></td></tr></table></td></tr></table>"
				end if
				end sub
				sub usererr()
				response.write "<table width=100% align=center border=0 cellspacing=0 cellpadding=0  bordercolor=#CCCCCC><tr><td bordercolor=#FFFFFF bgcolor=#FFFFFF align=center>"
				response.write "<table width=450 border=0 align=center cellpadding=2 cellspacing=0><tr><td height=260>"
				response.write "<p>· <font color=red>用户注册失败!</font></p><p>· 您输入的用户名或e-mail地址已存在，请返回重新输入!</p><p>· <a href=javascript:history.go(-1)>返回上一页</a></p> </td></tr></table></td></tr></table>"
				end sub
				%></TD>
        </TR>
      </TBODY>
    </TABLE></td>
  </tr>
  <tr><td height="40" colspan="2">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
