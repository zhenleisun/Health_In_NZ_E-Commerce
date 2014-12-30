<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--商品推荐</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="3" topmargin="4" marginwidth="0">
<%dim bookid,action
bookid=request.QueryString("id")
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from products where bookid="&bookid,conn,1,1
bookname=rs("bookname")
rs.close
set rs=nothing

set rs=server.CreateObject("adodb.recordset")
rs.Open "select mailaddress,mailusername,mailuserpass,mailname,mailsend from webinfo",conn,1,1
mailaddress=rs("mailaddress")
mailusername=rs("mailusername")
mailuserpass=rs("mailuserpass")
mailname=rs("mailname")
mailsend=rs("mailsend")
rs.close
set rs=nothing

'发邮件
	topic=webname &"　的商品推荐"
	mailbody="<html>"
	mailbody=mailbody & "<title>商品推荐</title>"
	mailbody=mailbody & "<body>"
	mailbody=mailbody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody & "<TD valign=middle align=top>"
	mailbody=mailbody & trim(request("friendname"))&"，您好：<br><br>"
	mailbody=mailbody & "您的朋友"&trim(request("myname"))&"推荐给您一款商品，请点击以下链接查看详细情况<br><br>"
	mailbody=mailbody & "商品名称：　<a href=http://"&weburl&"/list.asp?id="&bookid&" >"&bookname&"</a><br><br><br>"
	mailbody=mailbody & "欢迎光临：　<a href=http://"&weburl&" >"&webname&"</a><br><br><br>"
	mailbody=mailbody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	mailbody=mailbody & "</body>"
	mailbody=mailbody & "</html>"

	on error resume next
Set JMail=Server.CreateObject("JMail.Message")
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
jmail.from = mailsend
jmail.silent = true
jmail.Logging = true
jmail.FromName = mailname
jmail.mailserverusername = mailusername
jmail.mailserverpassword = mailuserpass
jmail.AddRecipient request("friendemail")
jmail.body=mailbody
JMail.Subject=topic
if not jmail.Send ( mailaddress ) then
SendMail=""
else
SendMail="OK"
end if



	'on error resume next
	'dim JMail
	'Set JMail=Server.CreateObject("JMail.SMTPMail")
	'JMail.Logging=True
	'JMail.Charset="gb2312"
	'JMail.ContentType = "text/html"
	'JMail.ServerAddress="61.145.114.64"
	'JMail.Sender=webemail
	'JMail.Subject=topic
	'JMail.Body=mailbody
	'JMail.AddRecipient request("friendemail")
	'JMail.Priority=1
	'JMail.Execute 
	'Set JMail=nothing 
	'if err then 
	'SendMail=err.description
	'err.clear
	'else
	'SendMail="OK"
	'end if

	'on error resume next
	'dim  objCDOMail
	'Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	'objCDOMail.From =webemail
	'objCDOMail.To =request("friendemail")
	'objCDOMail.Subject =topic
	'objCDOMail.BodyFormat = 0 
	'objCDOMail.MailFormat = 0 
	'objCDOMail.Body =mailbody
	'objCDOMail.Send
	'Set objCDOMail = Nothing
	'if err then 
	'SendMail=err.description
	'err.clear
	'else
	'SendMail="OK"
	'end if
	
	if SendMail="OK" then
		sendmsg="您的推荐已经成功发出！"
	else
		sendmsg="由于系统错误，您的推荐未能成功发送。"
	end if
	response.Write "<script language='javascript'>alert('"&sendmsg&"');window.close();</script>"
	end if
	%>
<table width="300" align="center" border="0" cellpadding="2" cellspacing="1" bgcolor="#cccccc">
                          <tr> 
                            <form name="pinglunform" method="post" action="shangpintj.asp?action=save&id=<%=bookid%>" onsubmit="return check();">
                              <td bgcolor="#ffffff"> 
                                <table width="100%" border="0" cellpadding="2" cellspacing="1">
                                  <tr align="center"> 
                                    <td colspan="2"><b>商　品　推　荐</b></td>
                                  </tr>
                                  <tr> 
                                    <td width="40%" align="right">朋友的姓名：</td>
                                    <td width="60%"> 
                                      <input class="wenbenkuang" name="friendname" type="text" id="friendname">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td align="right">朋友的e-mail：</td>
                                    <td> 
                                      <input class="wenbenkuang" name="friendemail" type="text" id="friendemail">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td align="right">您的大名：</td>
                                    <td> 
                                      <input class="wenbenkuang" name="myname" type="text" id="myname" value=<%=request.Cookies("cnhww")("username")%>>
									</td>
                                  </tr>
                                  <tr> 
                                    <td colspan="2" height="30" align="center">
									<input class="go-wenbenkuang" name="submit" value="开始发送" type="submit" onFocus="this.blur()">
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </form>
                          </tr>
                        </table>
</body>
</html>
<%function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function%>
<script LANGUAGE="javascript">
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.pinglunform.friendname.value)) {
	document.pinglunform.friendname.focus();
    alert("请填写朋友的姓名！");
	return false;
  }
  if(checkspace(document.pinglunform.friendemail.value)) {
	document.pinglunform.friendemail.focus();
    alert("请填写朋友的Email！");
	return false;
  }
 if(document.pinglunform.friendemail.value.length!=0)
  {
    if (document.pinglunform.friendemail.value.charAt(0)=="." ||        
         document.pinglunform.friendemail.value.charAt(0)=="@"||       
         document.pinglunform.friendemail.value.indexOf('@', 0) == -1 || 
         document.pinglunform.friendemail.value.indexOf('.', 0) == -1 || 
         document.pinglunform.friendemail.value.lastIndexOf("@")==document.pinglunform.friendemail.value.length-1 || 
         document.pinglunform.friendemail.value.lastIndexOf(".")==document.pinglunform.friendemail.value.length-1)
     {
      alert("请正确填写朋友的Email！");
      document.pinglunform.friendemail.focus();
      return false;
      }
   }
 else
  {
   alert("请填写朋友的Email");
   document.pinglunform.friendemail.focus();
   return false;
   }
  if(checkspace(document.pinglunform.myname.value)) {
	document.myname.mayname.focus();
    alert("请填写您的姓名！");
	return false;
  }
  }
</script>