<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/fahuo.gif" width="180" height="50"></td>
  </tr>
</table><marquee onmouseover=this.stop() onmouseout=this.start() direction="up" width="100%" height="85"  scrollamount="2" > 
<table width="100%" height="155"  border="0" cellpadding="0" cellspacing="0">

  <tr>
    <td height="124"  STYLE='PADDING-LEFT: 18px'> 
  <%

mess=split(fahuomess,"<br>")
    for i=0 to ubound(mess) %>
	
	<%=trim(mess(i))%>	<hr>
<%next%>
</table>
</marquee>

