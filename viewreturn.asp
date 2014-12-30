<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--意见反馈</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
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
        <td><!--#include file="searc.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width="812" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="28" background="images/body/pdbg01.gif"><div align="left">&nbsp;&nbsp;&nbsp;&nbsp;尊敬的客户，留下您宝贵的建议，我们愿为您提供更好的服务！</div></td>
      </tr>
      <tr>
        <td><span class="b">
          
        </span>
          <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
              <tr> 
                <td bordercolor="#FFFFFF" bgcolor="#FFFFFF"><br>
                  <table width="95%"  border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td valign="top"> <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td background="img/bg1_4color_.jpg"><img src="img/bg1_4color_.jpg" width="1" height="3"></td>
                          </tr>
                        </table>
                        <%if ly="0" then%>
                        <FORM action=fksave.asp method=post name="form" id="form">
                          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
                            <tr> 
                              <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                            </tr>
                            <tr> 
                              <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp; 
                                </strong><strong>用户留言</strong></td>
                            </tr>
                            <tr> 
                              <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                            </tr>
                          </table>
                          <br>
                          <TABLE width="90%" border=0 align="center" cellPadding=0 cellSpacing=1>
                            <TBODY>
                              <TR> 
                                <TD width="407" height="30">&nbsp;&nbsp; 姓名： 
                                  <input name=name class="wenbenkuang" id="name"></TD>
                              </TR>
                              <TR> 
                                <TD height="30"> &nbsp;&nbsp; OICQ： 
                                  <input name=qq class="wenbenkuang" id="qq">
                                  性别： 
                                  <label> 
                                  <input name="sex" type="radio" value="boy" checked>
                                  先生</label> <label> 
                                  <input type="radio" name="sex" value="girl">
                                  小姐</label></TD>
                              </TR>
                              <TR> 
                                <TD height="30">&nbsp;&nbsp; 网址： 
                                  <input name=url class="wenbenkuang" id="url" value="http://">
                                  邮箱： 
                                  <input  name=email class="wenbenkuang" id="email">                                </TD>
                              </TR>
                              <TR> 
                                <TD height="30">&nbsp;&nbsp; 留言内容：</TD>
                              </TR>
                              <TR> 
                                <TD height="130" noWrap> &nbsp;&nbsp; <textarea name=content cols=60 rows=8 id="content" class="wenbenkuang"></textarea>                                </TD>
                              <TR> 
                                <TD height="30" colspan="2"> &nbsp;&nbsp; <input name="submit2" type=submit class="go-wenbenkuang" id="submit4" value=我要留言> 
                                  &nbsp;&nbsp; <input name="submit32" type=reset class="go-wenbenkuang" id="submit32" value=清除重来> 
                                  <%
Randomize '初始代随机数种子
num1=rnd() '产生随机数num1
num1=int(26*num1)+65 '修改num1的范围以使其是A-Z范围的Ascii码，以防表单名出错
session("antry")="test"&chr(num1) '产生随机字符串
%>
                                  <input name="temp" type="hidden" id="temp" value="<%=session("antry")%>">                                </TD>
                              </TR>
                            </TBODY>
                          </TABLE>
                        </FORM>
                        <p>&nbsp; </p>
                        <p> 
                          <%else%>
                        </p>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
                                <tr> 
                                  <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                </tr>
                                <tr> 
                                  <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp; 
                                    </strong><b>留言信息</b></td>
                                </tr>
                                <tr> 
                                  <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                </tr>
                              </table>
                              <table width="98%" border="0" align="center">
                                <tr> 
                                  <td align="center"> 
                                    <%
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
	IF Err.Number <>0 Then
		unHtml= "HTML转换中出错请联系管理员<br>"
		Err.Clear
	End IF
End Function
%>
                                    <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from guestbook where online=1 order by id desc",conn,3,2
'dim page
'dim e_page
'e_page=5 '每页显示留言数
'rs.pagesize=e_page
'if request.querystring("page")="" or request.querystring("page")=0 then
'page=1
'else
'page=request.querystring("page")
'rs.absolutepage=trim(request.querystring("page"))
'end if
%>
                                    <SCRIPT language=JavaScript>

function is_number(str)
{
	exp=/[^0-9()-]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}

function CheckInput(){

	if(form.name.value==''){
		alert("您没有填写昵称！");
		form.name.focus();
		return false;
	}
	if(form.name.value.length>20){
		alert("昵称不能超过20个字符！");
		form.name.focus();
		return false;
	}

	if(!is_number(document.form.qq.value)){
		alert("QQ号必须是数字！");
		form.qq.focus();
		return false;
	}

	if(form.content.value==''){
		alert("您没有填写留言内容！");
		form.content.focus();
		return false;
	}
	if(form.content.value.length>255){
		alert("留言内容不能超过255个字符！");
		form.content.focus();
		return false;
	}
	
	return true;
}
</SCRIPT>
                                    <br> 
                                    <%
if rs.eof and rs.bof then
%>
                                    <table width="380" border="0" align="center" cellpadding="4" >
                                      <tr> 
                                        <td height="40" align="center"> <p>还没有任何留言！</p></td>
                                      </tr>
                                    </table>
                                    <%else%>
                                    <%
	  	rs.PageSize =5 '每页记录条数
		iCount=rs.RecordCount '记录总数
		iPageSize=rs.PageSize
    	maxpage=rs.PageCount 
    	page=request("page")
    	per_page=rs.PageSize
    
    if Not IsNumeric(page) or page="" then
        page=1
    else
        page=cint(page)
    end if
    
    if page<1 then
        page=1
    elseif  page>maxpage then
        page=maxpage
    end if
    
    rs.AbsolutePage=Page

	if page=maxpage then
		x=iCount-(maxpage-1)*iPageSize
	else
		x=iPageSize
	end if
%>
                                    <table width="100%" border="0">
                                      <tr> 
                                        <td colspan="12" height="25" align="center" bgcolor="#FFFFFF" > 
                                          <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>                                        </td>
                                      </tr>
                                    </table>
                                    <%
For i=1 To x
'do while not rs.eof and e_page>0
%>
                                    <TABLE width=100% border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
                                      <TBODY>
                                        <TR> 
                                          <TD bgcolor="f1f1f1"><TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                                              <TBODY>
                                                <TR bgcolor="f1f1f1"> 
                                                  <TD width="67" height="23" align="center" bgcolor="f1f1f1"><font color="#FF0000"><strong>第<%=rs("id")%>条</strong></font></TD>
                                                  <TD width="464" align="right"> 
                                                    <%if rs("qq")<>"" then%>
                                                    <img src="images/qq.gif" width="16" height="16" border="0"> 
                                                    <%=rs("qq")%> 
                                                    <%end if%>
                                                    &nbsp; 
                                                    <%if rs("email")<>"" then%>
                                                    <a href="mailto:<%=rs("email")%>" target="_blank" title=Email：<%=rs("email")%>><img src="images/email.gif" width="14" height="15" border="0"></a> 
                                                    <a href="mailto:<%=rs("email")%>" target="_blank" title=Email：<%=rs("email")%>><%=rs("email")%></a> 
                                                    <%end if%>
                                                    <%if rs("url")<>"" then%>
                                                    <a href="<%=rs("url")%>" target="_blank" title=主页：<%=rs("url")%>><img src="images/url.gif" width="16" height="16" border="0"></a> 
                                                    <a href="<%=rs("url")%>" target="_blank" title=主页：<%=rs("url")%>><%=rs("url")%></a> 
                                                    <%end if%>
                                                    &nbsp;</TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE></TD>
                                        </TR>
                                        <TR> 
                                          <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD bgcolor="#FFFFFF"> <TABLE width="100%" border=0 style="table-layout:fixed;word-break:break-all">
                                                      <TBODY>
                                                        <TR> 
                                                          <TD width="109"><img src="images/<%=rs("sex")%>.gif"></TD>
                                                          <TD colspan="2"> <strong><font color="<%=flzhbt%>">姓名：</font></strong><font color="<%=flzhzt%>"><%=rs("name")%></font><br> 
                                                            <strong><font color="<%=flzhbt%>">内容：</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("content"))%></font></TD>
                                                        </TR>
                                                        <TR> 
                                                          <TD colspan="2">&nbsp;</TD>
                                                          <TD width="250" align="right"><%=rs("time")%></TD>
                                                        </TR>
                                                      </TBODY>
                                                    </TABLE></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE></TD>
                                        </TR>
                                        <%if rs("reply")<>"" then%>
                                        <TR> 
                                          <TD bgColor=#f2f2f2> <table width="100%" border="0" style="table-layout:fixed;word-break:break-all">
                                              <tr> 
                                                <td width="10">&nbsp;</td>
                                                <td width="567"><font color="#FF0000">管理员回复：</font><br> 
                                                  <%=unHtml(rs("reply"))%></td>
                                              </tr>
                                            </table></TD>
                                        </TR>
                                        <%end if%>
                                      </TBODY>
                                    </TABLE>
                                    <br> 
                                    <%
'e_page=e_page-1
'rs.movenext
'loop
		RS.MoveNext
next
%>
                                    <table width="100%" border="0">
                                      <tr> 
                                        <td> 
                                          <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>                                        </td>
                                      </tr>
                                    </table>
                                    <%
rs.close
set rs=nothing
end if
%>                                  </td>
                                </tr>
                                <tr> 
                                  <td align="center"><FORM action=fksave.asp method=post name="form" id="form">
                                      <br>
                                      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
                                        <tr> 
                                          <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                        </tr>
                                        <tr> 
                                          <td height="25" bgcolor="efefef"> <strong>&nbsp;&nbsp;<img src="img/dot_03.gif" width="9" height="9" border="0"> 
                                            </strong><strong>用户留言</strong></td>
                                        </tr>
                                        <tr> 
                                          <td height="1" bgcolor="cccccc"><img src="img/dot_03.gif" width="9" height="1" border="0"></td>
                                        </tr>
                                      </table>
                                      <TABLE width="100%" border=0 cellPadding=0 cellSpacing=1>
                                        <TBODY>
                                        <TD height="30">&nbsp;&nbsp; 姓名： 
                                          <input name=name class="wenbenkuang" id="name"></TD>
                                        </TR>
                                        <TR> 
                                          <TD height="30"> &nbsp;&nbsp; OICQ： 
                                            <input name=qq class="wenbenkuang" id="qq3">
                                            性别： 
                                            <label> 
                                            <input name="sex" type="radio" value="boy" checked>
                                            先生</label> <label> 
                                            <input type="radio" name="sex" value="girl">
                                            小姐</label></TD>
                                        </TR>
                                        <TR> 
                                          <TD height="30">&nbsp;&nbsp; 网址： 
                                            <input name=url class="wenbenkuang" id="url" value="http://">
                                            邮箱： 
                                            <input  name=email class="wenbenkuang" id="email3">                                          </TD>
                                        </TR>
                                        <TR> 
                                          <TD height="30">&nbsp;&nbsp; 留言内容：</TD>
                                        </TR>
                                        <TR> 
                                          <TD height="130" noWrap> &nbsp;&nbsp; 
                                            <textarea name=content cols=80 rows=8 id="content" class="wenbenkuang"></textarea>                                          </TD>
                                        <TR> 
                                          <TD height="30" colspan="2"> &nbsp;&nbsp; 
                                            <input name="submit" type=submit class="go-wenbenkuang" id="submit" value=我要留言> 
                                            &nbsp;&nbsp; <input name="submit3" type=reset class="go-wenbenkuang" id="submit3" value=清除重来> 
                                            <%
Randomize '初始代随机数种子
num1=rnd() '产生随机数num1
num1=int(26*num1)+65 '修改num1的范围以使其是A-Z范围的Ascii码，以防表单名出错
session("antry")="test"&chr(num1) '产生随机字符串
%>
                                            <input name="temp" type="hidden" id="temp" value="<%=session("antry")%>">                                          </TD>
                                        </TR>
                                      </TABLE>
                                    </FORM></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table>
                        <%end if%>                      </td>
                    </tr>
                  </table>
                  <br>                </td>
              </tr>
            </table></td>
      </tr>
    </table></td>
    </tr>
</table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function checkfk()
{
   if(checkspace(document.fkinfo.fksubject.value)) {
	document.fkinfo.fksubject.focus();
    alert("您没有填写主题！");
	return false;
  }
   if(checkspace(document.fkinfo.fkusername.value)) {
	document.fkinfo.fkusername.focus();
    alert("请填写您的姓名！");
	return false;
  }
   if(checkspace(document.fkinfo.fklaizi.value)) {
	document.fkinfo.fklaizi.focus();
    alert("请填写您来自哪里！");
	return false;
  }
     if(checkspace(document.fkinfo.fkcontent.value)) {
	document.fkinfo.fkcontent.focus();
    alert("请填写反馈信息内容！");
	return false;
  }
  if(document.fkinfo.fkemail.value.length!=0)
  {
    if (document.fkinfo.fkemail.value.charAt(0)=="." ||        
         document.fkinfo.fkemail.value.charAt(0)=="@"||       
         document.fkinfo.fkemail.value.indexOf('@', 0) == -1 || 
         document.fkinfo.fkemail.value.indexOf('.', 0) == -1 || 
         document.fkinfo.fkemail.value.lastIndexOf("@")==document.fkinfo.fkemail.value.length-1 || 
         document.fkinfo.fkemail.value.lastIndexOf(".")==document.fkinfo.fkemail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.fkinfo.fkemail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.fkinfo.fkemail.focus();
   return false;
   }
}
//-->
</script> 
