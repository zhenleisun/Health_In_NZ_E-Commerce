<TABLE width="1002" border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR>
      <TD><img src="images/xxl_10.png" width="1002" height="55" border="0" usemap="#Map2" /></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
    <%set rs=server.CreateObject("adodb.recordset")
	rs.open "select Top "&tejia&" * from products where tejiabook=1 and zhuangtai=0 order by bookid desc ",conn,1,1
	if rs.eof and rs.bof then
	response.write "<center><br><font color=red size=2>对不起，暂无此类商品！</font></font>"
	'response.End
    else
	if not rs.eof then
	i=1
	do while not rs.eof
	%>
    <td valign="top"><table border="0" align="center" cellpadding="0" cellspacing="0" >
      <tr>
        <td align="center"><%if rs("bookpic")="" then 
		response.write "<div align=center><a href=list.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
		else%>
        <a href="Products.asp?id=<%=rs("bookid")%>" target="_blank"> <img src="<%=trim(rs("bookpic"))%>" width="294" height="246" border="0" align="absmiddle" /></a>
        <%end if%></td>
      </tr>
	  <tr>
          <td height="42" align="left" ><%response.write "<a href=products.asp?id="&rs("bookid")&" ><font color=#1b1b1b style=""font-size:14px; font-weight:bold;"">"
			if len(trim(rs("bookname")))>7 then
			response.write left(trim(rs("bookname")),7)&".."
			else
			response.write trim(rs("bookname"))
			end if
			response.write "</font></a>"
			%></td>
        </tr>
          <tr>
            <td height="38" align="center" style="background:url(images/titlebg.jpg) no-repeat center center; font-size:14px; font-family:'微软雅黑'; color:#FFF; font-weight:bold;">会员价：<font color="#FFFFFF"><%=rs("huiyuanjia")%></font>元 </td>
          </tr>
		  <tr>
            <td height="15" align="center"></td>
          </tr>
    </table>
    </td>
    <%if i mod 3 = 0 then%>
  </tr>
  <%end if
  rs.movenext
  i=i+1
  loop
  rs.close
  end if
  end if 
  %>
</table>
<map name="Map2" id="Map2">
  <area shape="rect" coords="906,18,993,43" href="class.asp?lx=tejia" />
</map>