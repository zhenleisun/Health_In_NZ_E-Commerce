<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--�̳�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="��Ȥ���Ϲ���ϵͳ,��Ȥ���Ϲ���ϵͳʱ�а�,��Ȥ����ϵͳ,���Ϲ���ϵͳ,����ϵͳ,��Ȥ����,�̳�Դ��,�����̵�,�����̵�ϵͳ,����ע��,��������,��ΰ����">
<meta name="keywords" content="��Ȥ���Ϲ���ϵͳ,��Ȥ���Ϲ���ϵͳʱ�а�,��Ȥ����ϵͳ,���Ϲ���ϵͳ,����ϵͳ,��Ȥ����,�̳�Դ��,�����̵�,�����̵�ϵͳ,����ע��,��������,��ΰ����">

<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="include/head.asp"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="39" valign="top"><table width="100%" height="184" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="right" valign="top"><img src="images/body/jiao.gif" width="15" height="17"></td>
      </tr>
      <tr>
        <td height="90"><div align="right"><img src="images/ttts.gif" width="26" height="110" /></div></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table></td>
    <td width="190"valign="top" style="BORDER-bottom:#FF67A0 1px solid;BORDER-left:#FF67A0 1px solid; BORDER-right:#FF67A0 1px solid;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12"></td>
      </tr>
      <tr>
        <td><!--#include file="searc.asp"--><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table></td>
    <td valign="top"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
      <tr>
        <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> �̳�����</td>
      </tr>
      <tr>
        <td valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td><img src="images/mingle/trend.gif" width="580" height="59"></td>
          </tr>
        </table>
          <br>
          <table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><img src="images/mingle/informtop.gif" width="580" height="35"></td>
            </tr>
          </table>
          <br>
          <%set rs=server.createobject("adodb.recordset")
		rs.open "select * from news  order by adddate desc",conn,1,1
		if rs.recordcount=0 then 
		%>
            <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
              <tr>
                <td align=center>��������</td>
              </tr>
            </table>
            <%
		else
	  		rs.PageSize =20 
			iCount=rs.RecordCount 
			iPageSize=rs.PageSize
    		maxpage=rs.PageCount 
    		page=request("page")
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
            <table width="580" border="0" align="center" cellpadding="2" cellspacing="1">
              <%
			For i=1 To x
			%>
              <tr>
                <td width="6%" height=24 align=center bgcolor="#FFFFFF" ><img src="images/body/dian3.gif"></td>
                <td  height=24 bgcolor="#FFFFFF" ><a href="trends.asp?id=<%=rs("newsid")%>"><%=trim(rs("newsname"))%></a>&nbsp<font color=red>(<%=rs("viewcount")%>)</font></td>
                <td width="19%" height=24 align=center bgcolor="#FFFFFF" ><%=year(rs("adddate"))&"."&month(rs("adddate"))&"."&day(rs("adddate"))%></td>
              </tr>
              <tr>
                <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
              </tr>
              <%
			rs.movenext
		    next
			%>
            </table>
            <%
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
Sub PageControl(iCount,pagecount,page,table_style,font_style)
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    Response.Write("<table width=90% border=0 cellpadding=0 cellspacing=0  align=center >" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=center height=35>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
    if page<=1 then
        Response.Write ("��ҳ " & vbCrLf)        
        Response.Write ("��ҳ " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">��ҳ</A> " & vbCrLf)
    end if
    if page>=pagecount then
        Response.Write ("��ҳ " & vbCrLf)
        Response.Write ("βҳ " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">βҳ</A> " & vbCrLf)            
    end if
    Response.Write(" ҳ�Σ�" & page & "/" & pageCount & "ҳ" &  vbCrLf)
    Response.Write(" ����" & iCount & "������" &  vbCrLf)
    Response.Write(" ת��" & "<INPUT TYEP=TEXT CLASS=wenbenkuang NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "ҳ"  & vbCrLf & "<INPUT CLASS=go-wenbenkuang type=submit value=GO>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>
        </td>
      </tr>
    </table></td>
    <td width="68"style="BORDER-left:#FF67A0 1px solid;">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
