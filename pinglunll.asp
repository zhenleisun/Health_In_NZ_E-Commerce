<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--User comments</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="5" marginwidth="0">
<%dim bookid,action
pinglunid=request.QueryString("id")%>
<table width="95%" align="center" border="0" cellpadding="2" cellspacing="1" bgcolor="#cccccc">
<tr> 
<td bgcolor="#ffffff"> 
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from  review where pinglunid="&pinglunid,conn,1,1
%>
<table width="95%" align="center" border="0" cellpadding="2" cellspacing="1">
<%set rs1=server.CreateObject("adodb.recordset")
rs1.open "select * from products where bookid="&rs("bookid"),conn,1,1
%>

<tr> 
<td> 
<table width="100%" border="0" cellpadding="2" cellspacing="1">
<tr> 
<td width="35%" align="right">评论商品：</td>
<td width="75%"><%=rs1("bookname")%></td>
</tr>
<tr>
  <td align="right">姓 名：</td>
  <td><%=rs("pinglunname")%> </td>
</tr>
<tr> 
<td width="35%" align="right">评 级：</td>
<td width="75%"> <img src="images/level<%=rs("pingji")%>.gif"></td>
</tr>
<tr> 
<td width="35%" align="right">评论标题：</td>
<td width="75%"> <%=rs("pingluntitle")%> </td>
</tr>
<tr> 
<td width="35%" align="right">评论正文：</td>
<td width="75%" style="table-layout:fixed;word-break:break-all"> 
<%=rs("pingluncontent")%> </td>
</tr>
<%if rs("huifu")<>"" then %>
<tr> 
<td width="35%" align="right"><font color="#FF0000">管理回复：</font></td>
<td width="75%" style="table-layout:fixed;word-break:break-all"> 
<%=rs("huifu")%> </td>
</tr>
<%end if%>
<tr>
<td width="35%"></td>
<td><input name="imageField" type="image" src="images/gb.gif" width="54" height="18" border="0" onFocus="this.blur()" onClick="javascript:window.close();"></td>
</tr>
</table>
<%rs.close
set rs=nothing%>
</td>
</tr>
</table>
</td>
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
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.pinglunform.pinglunname.value)) {
	document.pinglunform.pinglunname.focus();
    alert("请填写您的姓名！");
	return false;
  }
  if(checkspace(document.pinglunform.pingluntitle.value)) {
	document.pinglunform.pingluntitle.focus();
    alert("请填写评论标题！");
	return false;
  }
  if(checkspace(document.pinglunform.pingluncontent.value)) {
	document.pinglunform.pingluncontent.focus();
    alert("请填写评论正文！");
	return false;
  }
	  }
//-->
</script><%rs1.close
set rs1=nothing%>