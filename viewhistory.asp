<table width="190" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/viewhist.gif" width="190" height="50"></td>
  </tr>
  <tr>
    <td background="images/leftbg.jpg"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" >
	  <%set rs=server.createobject("adodb.recordset")
		ProductList = request.cookies("leisure")("fav")
		If ProductList<>"" then
		sql="Select * From products Where bookid In (" & ProductList & ")"
		rs.open sql,conn,1,1
		if rs.eof and  rs.bof then 
		response.write "<div align=center>您目前没有浏览过任何商品</div>"
		end if 
		do while not rs.eof 
		%>
		<tr>
		  <td width="21%" align="center" ><img src="images/ring01.gif"></td> 
		  <td width="79%" height="30" align="left" ><a href="products.asp?id=<%=rs("bookid")%>"> 
	      <%response.write trim(rs("bookname"))%></a></td>
	  </tr>
	  <%rs.movenext
		loop
		else
		response.write "<tr><td align=center>您目前没有浏览过任何商品</td></tr>"
		end if 
		%>
	</table></td>
  </tr>
</table>







		