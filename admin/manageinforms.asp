<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<body>

<%dim selectm,selectkey,selectbookid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if
selectbookid=request("selectbookid")
if selectbookid<>"" then
conn.execute "delete FROM inform where newsid in ("&selectbookid&")"
response.Redirect "manageinforms.asp"
response.End
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
</tr>
<tr> 
<form name="form1" method="post" action="">
<td height="100"> 
                                <%
				Const MaxPerPage=20
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages
   				dim j
   				dim sql
    				if Not isempty(request("page")) then
      				currentPage=Cint(request("page"))
   				else
      				currentPage=1
   				end if 
			set rs=server.CreateObject("adodb.recordset")
			select case selectm
			case ""
            rs.open "select newsid,title,updatetime FROM inform order by updatetime desc",conn,1,1
		    case "0"
			rs.open "select newsid,title,updatetime FROM inform order by updatetime desc",conn,1,1
		    case "newsid"
			rs.open "select newsid,title,updatetime FROM inform where newsid="&selectkey&" order by updatetime desc",conn,1,1
			case "title"
			rs.open "select newsid,title,updatetime FROM inform where title like '%"&selectkey&"%' order by updatetime desc",conn,1,1
			case "content"
			rs.open "select newsid,title,updatetime FROM inform where content like '%"&selectkey&"%' order by updatetime desc",conn,1,1
		  end select
		   	if err.number<>0 then
				response.write "数据库中无数据"
				end if
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
   				else
	  				totalPut=rs.recordcount
      				if currentpage<1 then
          				currentpage=1
      				end if
      				if (currentpage-1)*MaxPerPage>totalput then
	   					if (totalPut mod MaxPerPage)=0 then
	     					currentpage= totalPut \ MaxPerPage
	   					else
	      					currentpage= totalPut \ MaxPerPage + 1
	   					end if
      				end if
       				if currentPage=1 then
            			showContent
            			showpage totalput,MaxPerPage,"manageinforms.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"manageinforms.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"manageinforms.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0%>
                <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      				    <tr> 
      				      <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商城专题管理</font></b></td>
			      </tr>
                                  <tr > 
                                    <td width="10%" align="center" bgcolor="fbc2c2">序号</td>
                                    <td width="50%" align="center" bgcolor="fbc2c2">文章标题</td>
                                    <td width="30%" align="center" bgcolor="fbc2c2">加入时间</td>
                                    <td width="10%" align="center" bgcolor="fbc2c2">选 择</td>
                                  </tr>
                                  <%
		  do while not rs.eof%>
                                  <tr > 
                                    <td> 
                                      <div align="center"><a href=editbook.asp?id=<%=rs("newsid")%>> 
                                        </a><%=rs("newsid")%></div>
                                    </td>
                                    <td width="50%" STYLE='PADDING-LEFT: 10px'> 
                                      <a href=editinforms.asp?id=<%=rs("newsid")%>> 
                                      <% if len(replace(trim(rs("title")),"<br>",""))>16 then
			response.write left(replace(trim(rs("title")),"<br>",""),14)&"..."
			else
			response.write replace(trim(rs("title")),"<br>","")
			end if%>
                                      </a></td>
                                    <td align="center"><%=rs("updatetime")%></td>
                                    <td align="center">
									<input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("newsid")%>">
                                    </td>
                                  </tr>
                                  <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                                  <tr > 
                                    <td height="30" colspan="4" align="right">
									全选
									<input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
									<input type="submit" name="Submit" value="删 除" onClick="return test();">
									&nbsp;&nbsp;
									</td>
                                  </tr>
        </table>
                                <%  
				End Sub   
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				Response.Write "<form method=Post action="&filename&"?selectm="&selectm&"&selectkey="&selectkey&" >"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>上一页</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"专题 " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok' ></form>"  
				End Function  
			%>
                                <table width="12" height="7" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height=7></td>
                                  </tr>
                                </table>
      </td>
                            </form>
                          </tr>
                        </table>
					<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      				    <tr> 
      				      <td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">专题文章查询</font></b></td>
     				     </tr>
                          <tr> 
                            <form name="form2" method="post" action="manageinforms.asp">
                              <td height="50" > 
                                <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="35%" align="center">
									<input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关键字">
                                    </td>
                                    <td width="45%" align="center">
                                        <select name="selectm" id="selectm">
                                          <option value="0">全部文章</option>
                                          <option value="newsid">按文章序号</option>
                                          <option value="title">按文章标题</option>
                                          <option value="content">按文章内容</option>
                                        </select>
                                    </td>
                                    <td width="20%" height="30" align="center">
									<input type="submit" name="Submit2" value="查 询">
                                    </td>
                                  </tr>
								  <tr><td colspan=3 height=40>　除查询“所有商品”外，必须要输入关键字，才能查询<br>　根据商品序号来查询时，关键字里只能输入数字。</td></tr>
                                </table>
                              </td>
                            </form>
                          </tr>
                        </table>
<!--#include file="foot.asp"-->
</body>
</html>
<script>
function test()
{
  if(!confirm('确认删除吗？')) return false;
}
</script>
<script language=javascript>
function mm()
{
   var a = document.getElementsByTagName("input");
   if(a[0].checked==true){
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = false;
   }
   else
   {
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = true;
   }
}
</script>
