<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if%>
<html><head><title><%=webname%>--�û�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="5" marginwidth="0" >
<%dim bookid,action
pinglunid=request.QueryString("id")
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from review where pinglunid="&pinglunid,conn,1,3
rs("huifu")=HTMLEncode2(trim(request("huifu")))
rs("huifudate")=now()
rs.update
rs.close
set rs=nothing
response.write "<script language=javascript>alert('���Ļظ��ѳɹ��ύ����');history.go(-1);</script>"
response.End
end if
%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��������</font></b></td>
</tr>
<tr> 
<form name="pinglunform" method="post" action="review.asp?action=save&id=<%=pinglunid%>">
<td > 
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from review where pinglunid="&pinglunid,conn,1,3
%>
<table width="100%" align="center" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
<tr > 
<td width="23%" align="right">�� ����</td>
<td width="77%"> 
<input name="pinglunname" type="text" id="pinglunname" size="30" value="<%=rs("pinglunname")%>" readonly>
</td>
</tr>
<tr > 
<td align="right">�� ����</td>
<td><img src="../images/level<%=rs("pingji")%>.gif">
</td>
</tr>
<tr > 
<td align="right">���۱��⣺</td>
<td><input name="pingluntitle" type="text" id="pingluntitle" size="30" value="<%=rs("pingluntitle")%>" readonly></td>
</tr>
<tr > 
<td align="right" valign="top">�������ģ�</td>
<td><textarea name="pingluncontent" cols="30" rows="3" id="pingluncontent" readonly><%=rs("pingluncontent")%></textarea></td>
</tr>
<tr > 
<td align="right" valign="top"><font color="#FF0000">�����ظ���</font></td>
<td><textarea name="huifu" cols="30" rows="5" id="huifu"><%=rs("huifu")%></textarea></td>
</tr>
<tr >
<td></td>
<td>
<input onClick="return check();" name="submit" type="submit" value="�ظ�����">
<input onClick="ClearReset()" type="reset" name="Clear" value="������д"></td>
</tr>
</table>
<%rs.close
set rs=nothing%>
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
    alert("����д����������");
	return false;
  }
  if(checkspace(document.pinglunform.pingluntitle.value)) {
	document.pinglunform.pingluntitle.focus();
    alert("����д���۱��⣡");
	return false;
  }
  if(checkspace(document.pinglunform.pingluncontent.value)) {
	document.pinglunform.pingluncontent.focus();
    alert("����д�������ģ�");
	return false;
  }
	  }
//-->
</script>