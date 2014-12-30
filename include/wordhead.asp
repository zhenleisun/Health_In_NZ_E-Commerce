
<link href="images/nbase1.css" rel="stylesheet" type="text/css" />
<div class="top">
	<div class="top_t">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="26%" height="33" align="left" style="color:#FFFFFF; font-size:12px;"><a class="top_kuai" href="http://www.www.cn" onclick="window.external.addfavorite('http://www.www.cn','新西兰大型综合购物网！');return false">收藏纽之源</a>&nbsp;|&nbsp;<a class="top_kuai" href="Dingdan.asp">订单中心</a></td>
			<td width="54%">&nbsp;</td>
			<td width="20%" align="left" style="color:#FFFFFF; font-size:12px;"><a class="top_kuai" href="buy.asp?action=show">购物车</a>&nbsp;&nbsp;<a class="top_kuai" href="user.asp">登陆</a>&nbsp;|&nbsp;<a class="top_kuai" href="register.asp">注册</a>&nbsp;&nbsp;&nbsp;<a class="top_kuai" href="http://www.chinz56.co.nz/" target="_blank">快递查询</a></td>
		  </tr>
		</table>
	</div>
	<div class="top_c">
		<table width="1002" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="256" align="left"><img src="<%=weblogo%>" width="256" height="117" /></td>
			<td width="448"><img src="images/xxl_05.png" width="448" height="117" /></td>
			<td width="298" align="left" background="images/xxl_06.png">&nbsp;</td>
		  </tr>
	  </table>
	</div>
    <div class="top_nav">
	<table width="1002" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="8" height="15"></td>
      </tr>
      <tr>
        <td width="92" height="38" align="center"><a href="index.asp" class="nav">主页</a></td>
        <td width="98" align="center"><a href="class.asp?lx=big&anid=62" class="nav">蜂产品</a></td>
        <td width="110" align="center"><a href="class.asp?lx=big&anid=63" class="nav">保健品</a></td>
        <td width="106" align="center"><a href="class.asp?lx=big&anid=64" class="nav">母婴用品</a></td>
        <td width="126" align="center"><a href="class.asp?lx=big&anid=65" class="nav">护肤洗护</a></td>
        <td width="114" align="center"><a href="class.asp?lx=big&anid=67" class="nav">Happy tao</a></td>
        <td width="109" align="center"><a href="class.asp?lx=big&anid=66" class="nav">水果海鲜</a></td>
        <td width="247" align="right"><TABLE width="100%" border=0 align="right" cellPadding="1" cellSpacing="1">
                    <FORM action="research.asp" method="post" name="form2">
                      <TBODY>
                        <TR>
                          <TD width="39%" align=left style="padding-top:2px;">
                        <select name="anclassid" class="shortselect"><%set rs=server.CreateObject("adodb.recordset")
						  rs.open "select * from bsort order by anclassidorder",conn,1,1
						  %>
                              <option value="0">所有分类</option>
                              <%do while not rs.eof%>
                              <option value="<%=rs("anclassid")%>"> <%if len(trim(rs("anclass")))>4 then
								response.write left(trim(rs("anclass")),4)&"."
								else
								response.write trim(rs("anclass"))
								end if%></option>
                              <%rs.movenext
							  loop
							  rs.close
							  set rs=nothing%>
                   		  </select></TD>
						  <TD width="44%" style="padding-top:2px;"><INPUT class="wenbenkuang" name="searchkey" value="请输入产品名称" onfocus="this.select();" onclick="this.value = ''" size="14" maxlength="30" style=" border:none; background:none;" /></TD>
                          <TD width="17%" height="25" align="left"><INPUT name="Submit" type="submit" style="font-family:'微软雅黑'; width:18px; height:16px; background:url(images/sou.png); border:none;cursor:pointer;" value=""></TD>
                        </TR>
                    </FORM>
        </TABLE></td>
      </tr>
    </table>
	</div>
	<div class="top_e">
		<%
		set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from advertisement",conn,1,1
		tupian1=trim(rs("pic1"))
		tupian1url=trim(rs("pic1_lnk"))
		piaos1=trim(rs("tit1"))
		
		tupian2=trim(rs("pic2"))
		tupian2url=trim(rs("pic2_lnk"))
		piaos2=trim(rs("tit2"))
		
		tupian3=trim(rs("pic3"))
		tupian3url=trim(rs("pic3_lnk"))
		piaos3=trim(rs("tit3"))
		
		tupian4=trim(rs("pic4"))
		tupian4url=trim(rs("pic4_lnk"))
		piaos4=trim(rs("tit4"))
		 
		tupian5=trim(rs("pic5"))
		tupian5url=trim(rs("pic5_lnk"))
		 
		
		rs.Close
		set rs=nothing
		%>
		<script type="text/javascript">
		imgUrl1="<%=tupian1%>";
		imgtext1=""
		imgLink1=escape("<%=tupian1url%>");
		imgUrl2="<%=tupian2%>";
		imgtext2=""
		imgLink2=escape("<%=tupian2url%>");
		imgUrl3="<%=tupian3%>";
		imgtext3=""
		imgLink3=escape("<%=tupian3url%>");
		imgUrl4="<%=tupian4%>";
		imgtext4=""
		imgLink4=escape("<%=tupian4url%>");
		imgUrl5="<%=tupian5%>";
		imgtext5=""
		imgLink5=escape("<%=tupian5url%>");
		
		 var focus_width=1002
		 var focus_height=423
		 var text_height=0
		 var swf_height = focus_height+text_height
		 
		 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
		 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
		 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
		 
		 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
		 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/focus2.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
		 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
		 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
		 document.write('<embed src="images/focus2.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="1002" height="423" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"></embed>');
		 document.write('</object>');
		</script>
	</div>
	
</div>