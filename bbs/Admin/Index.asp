<!--#include file="Check.Asp"-->
<%
Dim Action
Action=Request.QueryString("Action")

Select Case Action
	Case"top"
		top()
	Case"left"
		Admin_left()
	Case Else
		Main()
End select
Set YxBBs=Nothing

Sub Main()
%>
<frameset rows="47,*" cols="*" frameborder="no" border="0" framespacing="0">
  <frame src="?action=top" name="topFrame" scrolling="No" noresize="noresize" id="topFrame" title="topFrame" />
  <frameset cols="200,*" frameborder="no" border="0" framespacing="0">
    <frame src="?action=left" name="leftFrame" noresize="noresize" id="leftFrame" title="leftFrame" />
    <frame src="Basic.Asp" name="main" id="main" title="mainFrame" />
  </frameset>
</frameset>
<noframes>


<%
end sub
sub top()
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<body>
<div style="width:1024px;BACKGROUND-IMAGE: url(../images/admin_top_bg.gif);">
<div style="width:200px;float: left;"><img src="../images/admin_top.jpg" /></div>
<div class="w291" style="float: left;">&nbsp;&nbsp;<p><b><font color="#ffffff">YxBBs 2.3 For Access��</font></b></p></div>
<div class="w390" style="float: left;">&nbsp;&nbsp;<p><script SRC="http://yimxu.com/js/ver3.js"></script></p></div>
<div style="clear: both;"></div> 
</div>
</body>


<%
End Sub
sub Admin_left()
%>
<script language="JavaScript">
var iever= new String("Microsoft Internet Explorer");
var bver=navigator.appName;
if(bver==iever){}
</script>
<body>
<div class="ta w160">
<div class="th jz">��̳�����˵�</div>
<div class="td jz w152 h50">[<a href="index.asp" target="_top">������ҳ</a>]��[<a href="login.asp?action=exit" target="_top">�˳�����</a>]<br />[<a href="javascript:void(0);" onclick="bbs_fun_zkbar();">չ������</a>]��[<a href="javascript:void(0);" onclick="bbs_fun_ssbar();">��������</a>]</div>
<div style="clear: both;"></div> 
</div>
<%call space()%>
<div class="ta w160">
<div class="th jz" onclick="javascript:bbs_fun_objdisplay(document.all.leftmenumain1);" style="cursor:hand">��̳��������</div>
<div class="td jz w152 h120" id="leftmenumain1"><small>��</small> <a href="basic.asp?action=config" target="main">��̳��Ϣ����</a><br /><small>��</small> <a href="basic.asp?action=link" target="main">��̳���˹���</a><br /><small>��</small> <a href="basic.asp?action=lockip" target="main">i p ��������</a><br /><small>��</small> <a href="basic.asp?action=setad" target="main">���������</a><br /><small>��</small> <a href="basic.asp?action=updatebbs" target="main">��̳�����޸�</a></div>
<div style="clear: both;"></div> 
</div>
<%call space()%>
<div class="ta w160">
<div class="th jz" onclick="javascript:bbs_fun_objdisplay(document.all.leftmenumain2);" style="cursor:hand">�����������</div>
<div class="td jz w152 h120" id="leftmenumain2"><small>��</small> <a href="board.asp" target="main">��̳�������</a><br /><small>��</small> <a href="data.asp?action=delessay" target="main">������������</a><br /><small>��</small> <a href="data.asp?action=delsms" target="main">����ɾ������</a><br /><small>��</small> <a href="data.asp?action=allsms" target="main">Ⱥ���ʼ�����</a><br /><small>��</small> <a href="data.asp?action=recycle" target="main">���ӻ���ϵͳ</a></div>
<div style="clear: both;"></div> 
</div>
<%call space()%>
<div class="ta w160">
<div class="th jz" onclick="javascript:bbs_fun_objdisplay(document.all.leftmenumain3);" style="cursor:hand">��̳�û�����</div>
<div class="td jz w152 h95" id="leftmenumain3"><small>��</small> <a href="users.asp" target="main">�û���Ϣ����</a> | <a href="users.asp?action=isdel&order=3" target="main">���</a><br /><small>��</small> <a href="users.asp?action=addadmin" target="main">�����ʺ�����</a> | <a href="users.asp?action=admin" target="main">����</a><br /><small>��</small> <a href="users.asp?action=classadd" target="main">�û���������</a> | <a href="users.asp?action=userclass" target="main">����</a><br /><small>��</small> <a href="users.asp?action=addgrade" target="main">�û���������</a> | <a href="users.asp?action=usergrade" target="main">����</a></div>
<div style="clear: both;"></div> 
</div>

<%call space()%>
<div class="ta w160">
<div class="th jz" onclick="javascript:bbs_fun_objdisplay(document.all.leftmenumain4);" style="cursor:hand">���������</div>
<div class="td jz w152 h70" id="leftmenumain4"><small>��</small> <a href="plus.asp?action=plus" target="main">��̳�������</a><br /> <small>��</small> <a href="plus.asp?action=bank" target="main">��̨���й���</a><br /><small>��</small> <a href="plus.asp?action=skin" target="main">���������</a></div>
<div style="clear: both;"></div> 
</div>
<%call space()%>
<div class="ta w160">
<div class="th jz" onclick="javascript:bbs_fun_objdisplay(document.all.leftmenumain5);" style="cursor:hand">�����ļ�����</div>
<div class="td jz w152 h95" id="leftmenumain5"><small>��</small> <a href="data.asp" target="main">���ݱ�����</a>&nbsp;&nbsp;&nbsp;<br /><small>��</small> <a href="data.asp?action=executesql" target="main">ִ��sql���</a>&nbsp;&nbsp;<br /><small>��</small> <a href="data.asp?action=uploadfile" target="main">�ϴ��ļ�����</a></div>
<div style="clear: both;"></div> 
</div>
<%call space()%>
<div class="ta w160">
<div class="th jz">�����Ȩ��Ϣ</div>
<div class="td w152 jz h50"><small>��</small> ��Ȩ���У�&nbsp;&nbsp;<a target="_blank" href="http://www.yimxu.com">��  ��</a>&nbsp;&nbsp;<br /> <small>��</small> ����֧�֣�&nbsp; <a target="_blank" href="http://bbs.yimxu.com/">Yxbbs</a></div>
<div style="clear: both;"></div> 
</div>
</body>

<script language=javascript>
function bbs_fun_objzk(obj)
{
	if(bver==iever){obj.style.display = "block";}
}
function bbs_fun_objss(obj)
{
	if(bver==iever){obj.style.display = "none";}
}
function bbs_fun_objdisplay(obj)
{
	if(bver==iever){
		if(obj.style.display == "none")
		{
			obj.style.display = "block";
		}
		else
		{
			obj.style.display = "none";
		}
	}
}
function bbs_fun_zkbar()
{
	if(bver==iever){
		bbs_fun_objzk(document.all.leftmenumain1);
		bbs_fun_objzk(document.all.leftmenumain2);
		bbs_fun_objzk(document.all.leftmenumain3);
		bbs_fun_objzk(document.all.leftmenumain4);
		bbs_fun_objzk(document.all.leftmenumain5);
	}
}
function bbs_fun_ssbar()
{
	if(bver==iever){
		bbs_fun_objss(document.all.leftmenumain1);
		bbs_fun_objss(document.all.leftmenumain2);
		bbs_fun_objss(document.all.leftmenumain3);
		bbs_fun_objss(document.all.leftmenumain4);
		bbs_fun_objss(document.all.leftmenumain5);

	}
}
if(bver==iever){
	bbs_fun_objdisplay(document.all.leftmenumain1);
	bbs_fun_objdisplay(document.all.leftmenumain2);
	bbs_fun_objdisplay(document.all.leftmenumain3);
	bbs_fun_objdisplay(document.all.leftmenumain4);
	bbs_fun_objdisplay(document.all.leftmenumain5);

}
</script>
<%end sub%>
</html>