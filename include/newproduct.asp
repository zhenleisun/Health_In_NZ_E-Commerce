<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/np.gif" width="100%" height="30" border="0" usemap="#Map4"></td>
  </tr>
  <tr>
    <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" >
      <tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select Top "&xp&" * from products where newsbook=1 and zhuangtai=0 order by bookid desc ",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "<center><br><font color=red size=2>对不起，暂无此类商品！</font></font>"
		  'response.End
		  else
		  %>
        <%
		if not rs.eof then
		i=1
		do while not rs.eof%>
        <td height="110" valign="top" ><table width="128" align="center" cellpadding="0" cellspacing="0" >
              <tr valign="middle"> 
                <td></td>
              </tr>
              <tr> 
                <td height="113" align="center"><table width="98" height="100" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#bbbbbb" onmouseover="this.style.backgroundColor='#A10000'" onmouseout="this.style.backgroundColor=''">
                  <tbody>
                    <tr>
                      <td width="92" height="100" bgcolor="#ffffff" align="center"><%if rs("bookpic")="" then 
	response.write "<div align=center><a href=list.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
	else%>
                          <a href="Products.asp?id=<%=rs("bookid")%>" target="_blank"> <img src="<%=trim(rs("bookpic"))%>" width="115" height="115" border="0" align="absmiddle" /></a>
                          <%end if%>
                      </td>
                    </tr>
                  </tbody>
                </table></td>
              </tr>
              <tr> 
                <td align="center"><table width="95%" height="60" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="203" align="center" >
                        <%response.write "<a class=a4 href=products.asp?id="&rs("bookid")&" ><font color=#FF0000>"
			if len(trim(rs("bookname")))>7 then
			response.write left(trim(rs("bookname")),7)&".."
			else
			response.write trim(rs("bookname"))
			end if
			response.write "</font></a>"
			%>
                        <br /> </td>
                    </tr>
                    <tr> 
                      <td align="center" >市场价：<s><%=rs("shichangjia")&""%></s>元</td>
                    </tr>
                    <tr> 
                      <td align="center" >会员价：<font color="#FF0000"><%=rs("huiyuanjia")%></font>元 
                      </td>
                    </tr>
                    <tr> 
                      <td height="1" background="images/body/lineld.gif"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        <%if i mod 5 = 0 then%>
      </tr>
      <%end if
	    rs.movenext
         i=i+1
    	 loop
		rs.close
		end if
		end if 
	%>
    </table></td>
  </tr>
</table>


<map name="Map4" id="Map4">
  <area shape="rect" coords="68,7,80,15" href="class.asp?lx=news" />
</map>