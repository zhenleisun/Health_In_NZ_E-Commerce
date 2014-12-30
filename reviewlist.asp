<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--商品评论</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
              </tr>
    </table></td>
    <td width="812" valign="top" bgcolor="#FFFFFF"><%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""非法访问!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""非法访问!"");location.href=""index.asp"";</script>"
end if%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <%
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from products where bookid="&request("id"),conn,1,3
		if rs.recordcount>0 then
		spmx=rs("bookname")
		end if%>
        <tr>
          <td colspan="3" background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> 商品评论：<font color="ff0000"><%=spmx%></font></td>
        </tr>
      </table>
      <br>
      <table width="95%" align="center" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
        <tr>
          <td width="100%" valign="top"  align="center"><table width=100% align=center  border=0 cellpadding=0 cellspacing=0 bgcolor="#6FABCB">
                <tr>
                  <td align=center width=20%><font color="ffffff">主 题</font></td>
                  <td align=center width=60%><font color="ffffff">内 容</font></td>
                  <td align=center width=20%><font color="ffffff">评 价</font></td>
                </tr>
              </table>
            <br>
              <%set rs=server.createobject("adodb.recordset")
		rs.open "select * from review where bookid="&request("id")&" and shenhe=1 order by pinglunid desc",conn,1,1
		if rs.recordcount=0 then 
		%>
              <table width="370" border="0" cellspacing="0" cellpadding="5" align="center">
                <tr>
                  <td align=center>暂无评论</td>
                </tr>
              </table>
            <%
		else
	  		rs.PageSize =6'每页记录条数
			iCount=rs.RecordCount '记录总数
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
            <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center">
                  <%
			For i=1 To x
			%>
                  <tr>
                    <td bgcolor="#ffffff"><table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#e1e1e1">
                        <tr>
                          <td width="80%" colspan="2" bgcolor="#FFFFFF"><b>主 题：</b><%=trim(rs("pingluntitle"))%></td>
                          <td width="20%" bgcolor="#FFFFFF"><img src="images/level<%=rs("pingji")%>.gif"></td>
                        </tr>
                        <tr>
                          <td colspan="2" bgcolor="#FFFFFF"><b>作 者：</b><%=trim(rs("pinglunname"))%>　(<%=trim(rs("pinglundate"))%>)</td>
                          <td bgcolor="#FFFFFF"><%=trim(rs("ip"))%></td>
                        </tr>
                        <tr>
                          <td colspan="3" bgcolor="#FFFFFF"><b>内 容：</b>
                              <%if len(rs("pingluncontent"))>24 then 
							  response.write left(rs("pingluncontent"),20)&"...<a href=""pinglunll.asp?id="&rs("pinglunid")&""" title="&trim(rs("pingluntitle"))&">[详细内容]</a>" 
							  else response.write rs("pingluncontent")
							  end if%>                          </td>
                        </tr>
                        <%if trim(rs("huifu"))<>"" then%>
                        <tr>
                          <td colspan="3" bgcolor="#FFFFFF"><b><font color="#CC3333">回 复：</font></b><%=rs("huifu")%> </td>
                        </tr>
                        <%end if%>
                    </table></td>
                  </tr>
                  <%
			rs.movenext
		    next
			%>
            </table>
			
            <br>
              <!--#include file="reviews.asp"-->
              <%
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
Sub PageControl(iCount,pagecount,page,table_style,font_style)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    Response.Write("<table " & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=right>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
    if page<=1 then
        Response.Write ("首页 " & vbCrLf)        
        Response.Write ("上页 " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上页</A> " & vbCrLf)
    end if
    if page>=pagecount then
        Response.Write ("下页 " & vbCrLf)
        Response.Write ("尾页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">下页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">尾页</A> " & vbCrLf)            
    end if
    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)
    Response.Write(" 共有" & iCount & "种商品" &  vbCrLf)
    Response.Write(" 转到" & "<INPUT TYEP=TEXT NAME=page SIZE=1 Maxlength=5 VALUE=" & page & ">" & "页"  & vbCrLf & "<INPUT type=submit style=""font-size: 9pt"" value=GO class=b2>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>          </td>
        </tr>
		<tr><td height="30">&nbsp;</td>
  		</tr>
      </table></td>
    </tr>
	<tr><td height="40" colspan="2">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
