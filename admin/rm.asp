<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="css.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value
window.close();
</SCRIPT>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</HEAD>
<BODY>
<form name=rm>
&nbsp;&nbsp;文件地址：<input id=a name=a value='http://' style='width:200' title='可直接输入网上RealPlay文件的地址，或由上传程序自动产生RealPlay文件的地址' size="20"><br>
&nbsp;&nbsp;插入网上RealPlay文件支持格式为：rm、ra、ram<br>
&nbsp;&nbsp;文件注释：<input id=g name=g style='width:200' title='该注释将自动显示在文件的下方' size="20"><br>
&nbsp;&nbsp;立即播放：<select id=b><option value="-1">是<option value="0">否</select>
&nbsp;&nbsp;循环播放：<select id=c><option value="-1">是<option value="0">否</select><br>
&nbsp;&nbsp;文件位置：<select id=d><option value="left">居左<option value="center">居中<option value="right">居右</select>
&nbsp;&nbsp;显示选择：<select id=h><option value="0">只是音乐<option value="1">显示视频</select><br>
&nbsp;&nbsp;播放高度：<input id=e value='352' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" size="20">&nbsp;&nbsp;播放宽度：<input id=f value='240' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" size="20">
<br>&nbsp;&nbsp;<input type=button value='插入' id=Ok title='点击“插入”按钮，在编辑器中插入该RealPlay文件'></form>