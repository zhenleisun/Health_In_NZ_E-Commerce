<%

if webbanner<>"" then
QQ=split(webbanner,",")
for N=0 to UBound(QQ)
MyQQ=MyQQ+QQ(N)+":"
next
%>



<script>
var online= new Array();
if (!document.layers)
document.write('<div id="divStayTopright" style="position:absolute;">')
</script>
<layer id="divStayTopright">

<table border="0" width="73" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="73"><img border=0 src=images/mingle/sertop.gif width="73"></td>
  </tr>
  
  <tr> 
    <td valign=right  background=images/mingle/qqmiddle.gif >
	<%for i=0 to ubound(qq)%><a target=blank href=tencent://message/?uin=<%=qq(i)%>&Site=<%=webname%>&Menu=yes>&nbsp;<img src="http://wpa.qq.com/pa?p=1:<%=qq(i)%>:10" border="0" align="middle" ></a><br><%next%>
    </td> 
  </tr><tr><td height=3 background=images/mingle/qqmiddle.gif width=73 ></td></tr>
  <tr> 
    <td><img src="images/mingle/qqdown.gif" width=73></td>
  </tr>
  <tr> 
    <td background="images/mingle/serly.gif" width=73><a href=viewreturn.asp><img  src="images/mingle/serly.gif"border=0></a></td>
  </tr>
 
</table>
</layer>
<script type="text/javascript">
//Enter "frombottom" or "fromtop"
var verticalpos="frombottom"
if (!document.layers)
document.write('</div>')
function JSFX_FloatTopDiv()
{

	var startX =0,

	startY = 460;
	var ns = (navigator.appName.indexOf("Netscape") != -1);
	var d = document;
	function ml(id)
	{
		var el=d.getElementById?d.getElementById(id):d.all?d.all[id]:d.layers[id];
		if(d.layers)el.style=el;
		el.sP=function(x,y){this.style.right=x;this.style.top=y;};
		el.x = startX;
		if (verticalpos=="fromtop")
		el.y = startY;
		else{
		el.y = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
		el.y -= startY;
		}
		return el;
	}
	window.stayTopright=function()
	{
		if (verticalpos=="fromtop"){
		var pY = ns ? pageYOffset : document.body.scrollTop;
		ftlObj.y += (pY + startY - ftlObj.y)/8;
		}
		else{
		var pY = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
		ftlObj.y += (pY - startY - ftlObj.y)/8;
		}
		ftlObj.sP(ftlObj.x, ftlObj.y);
		setTimeout("stayTopright()", 10);
	}
	ftlObj = ml("divStayTopright");
	stayTopright();
}
JSFX_FloatTopDiv();
</script>
<% end if %>

