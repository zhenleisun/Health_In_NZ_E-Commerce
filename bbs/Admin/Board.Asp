<!--#include file="check.asp"-->
<%
const maxdepth=3
server.scripttimeout =999999
response.write("<body>")
call showtable("��̳�������","<center><a href='?'>��̳����</a> | <a href=""?action=addclass"">���ӷ���</a> | <a href=""?action=addboard"">������̳</a> | <a href=""?action=classorders"">��������</a> | <a href=""?action=boardunite"">��̳�ϲ�</a></center>")

response.flush
select case request("action")
case "addclass"
	addclass
case "saveclass"
	saveclass
case "editclass"
	editclass
case "saveeditclass"
	saveeditclass
case "delclass"
	delclass
case "addboard"
	addboard
case "saveboard"
	saveboard
case "editboard"
	editboard
case "saveeditboard"
	saveeditboard
case "delboard"
	delboard
case "classorders"
	classorders
case "saveclassorders"
	saveclassorders
case "boardunite"
	boardunite
case "saveboardunite"
	saveboardunite
case "passuser"
	passuser
case "savepassuser"
	savepassuser
case "ordersboard"
	ordersboard
case else
	boardinfo()
end select
adminfooter()

sub boardinfo
	dim rs,brs,i,install,temp,ii,po,boardtype
%>
<div class="ta">
<div class="th1 jz w390">��̳����</div>
<div class="th1 jz w390">��Ӧ����</div>

	<%
	set rs=yxbbs.execute("select boardid,boardname,parentid,depth,child,boardtype from [yx_board] order by rootid,orders")
	if rs.eof or rs.bof then
		call goback("","��̳û�з��࣡����<a href=board.asp?action=addclass> ���ӷ���</a>")
		exit sub
	end if
	brs=rs.getrows(-1)
	rs.close
	for i=0 to ubound(brs,2)
	temp=""
	install="&nbsp;&nbsp;<a href='?action=addboard&boardid="&brs(0,i)&"'>������̳</a>"
	if brs(3,i)=0 then'����
		temp="<div class=""td1 jz"" style=""text-align: left;"">"&temp
		if brs(4,i)>0 then'���������̳
			temp=temp&brs(1,i)&" ("&brs(4,i)&")"
		else
			temp=temp&brs(1,i)
		end if
		install=install & "&nbsp;|&nbsp;<a href='?action=editclass&boardid="&brs(0,i)&"'>�������</a>"
		if brs(4,i)>0 then
			install=install & "&nbsp;|&nbsp;<a href=javascript:alert('����ɾ�����÷��ຬ����̳!\n\nҪɾ�����࣬�����Ȱ����µ���̳ɾ�������ߡ�')>ɾ������</a>"
		else
			install=install & "&nbsp;|&nbsp;<a onclick=checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����') href='?action=delclass&boardid="&brs(0,i)&"'>ɾ������</a>"
		end if
	else'����
		temp="<div class=""td1 jz"" style=""text-align: left;"">"&temp
		po=""
			for ii=1 to brs(3,i)
				po=po&" <font color=red>�O</font> "
			next
		if brs(4,i)>0 then'���������̳
			temp=temp&po&brs(1,i)&" ("&brs(4,i)&")"
		else
			temp=temp&po&brs(1,i)
		end if
		install=install & " | <a href='?action=editboard&boardid="&brs(0,i)&"'>�޸�</a>"
		if brs(4,i)>0 then
			install=install & " | <a href='javascript:' onclick=alert('����ɾ�����ð��溬������̳!\n\nҪɾ�����棬�����Ȱ����µ�����̳ɾ�������ߡ�')>ɾ��</a>"
		else
			install=install & " | <a onclick=checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����') href='?action=delboard&boardid="&brs(0,i)&"'>ɾ��</a>"
		end if
		install=install & " | <a href='?action=ordersboard&boardid="&brs(0,i)&"'>��ǰ</a>"
		if brs(5,i) then
			install=install & " | <a href='?action=passuser&boardid="&brs(0,i)&"'>��֤�û�</a>"
		end if
	end if
	response.write(temp&"</div><div class=""td2 jz"" style=""text-align: left;"">"&install&"</div>")
next
%>

<div style='clear: both;'></div> 
</div>

<%
end sub

sub addclass
	dim newboardid
	set rs=yxbbs.execute("select max(boardid) from [yx_board]")
	if rs.eof or rs.bof then
		newboardid=1
	else
		newboardid=rs(0)+1
	end if
	if not isnumeric(newboardid) then newboardid=1
	rs.close
	%><form method=post style="margin:0" action="?action=saveclass">
        <div class="ta">
        <div class="th jz">���ӷ���</div>
        <div class="td1 h30"><b>�������ƣ�</b><br />��̳�ķ�������</div>
        <div class="td2 h30" style="text-align: left;"><input name="newboardid" type="hidden" value="<%=newboardid%>"><input type=text name="boardname" size=30></div>
        <div style="clear: both;"></div><div class="tf jz"><input type="submit" value=" �� �� ">&nbsp;&nbsp;<input type="reset" value=" �� �� "></div></form> </div>
	<%
end sub

sub saveclass
	dim boardname,newboardid,maxrootid
	boardname=yxbbs.fun.getstr("boardname")
	newboardid=yxbbs.fun.getstr("newboardid")
	if boardname="" or not isnumeric(newboardid) then
		call goback("",""):exit sub
	else
		set rs=yxbbs.execute("select boardid from [yx_board] where boardid="&newboardid)
		if not (rs.eof and rs.bof) then
			call goback("������ʾ","����ָ���ͱ����̳һ������ţ�")
			exit sub
		end if
		rs.close
		set rs=yxbbs.execute("select max(rootid) from yx_board")
		maxrootid=rs(0)+1
		if isnull(maxrootid) then maxrootid=1
		rs.close
		yxbbs.execute("insert into [yx_board](boardname,boardid,boardsetting,rootid,depth,parentid,orders,child,parentstr)values('"&boardname&"',"&newboardid&",'0,0,30,30720,10,0,0,0,0,0,0',"&maxrootid&",0,0,0,0,'0')")
		call suc("","�ɹ�����������̳���� <b>"&boardname&"</b> !","?")
	end if
end sub

sub editclass
	dim boardid
	set rs=yxbbs.execute("select boardname from[yx_board] where boardid="&yxbbs.boardid&"")
	if rs.eof then
		call goback("ϵͳ����","��̳�Ҳ���������࣬�����Ѿ�ɾ���ˡ�"):exit sub
	end if
	%><form method=post style="margin:0" action="?action=saveeditclass">
          <div class="ta">
        <div class="th jz">���ӷ���</div>
        <div class="td1 h30"><b>�������ƣ�</b><br />��̳�ķ�������</div>
        <div class="td2 h30" style="text-align: left;"><input name="boardid" type="hidden" value="<%=yxbbs.boardid%>"><input type=text name="boardname" value="<%=rs(0)%>" size=30></div>
        <div style="clear: both;"></div><div class="tf"><input type="submit" value=" �� �� ">&nbsp;&nbsp;<input type="reset" value=" �� �� "></div></form></div>
	<%rs.close
end sub

sub saveeditclass
	dim boardname,boardid
	boardname=yxbbs.fun.getstr("boardname")
	boardid=yxbbs.fun.getstr("boardid")
	if boardname="" or not isnumeric(boardid) then
		call goback("",""):exit sub
	else
		if yxbbs.execute("select boardid from [yx_board] where boardid="&boardid).eof then
			call goback("ϵͳ����","��̳�Ҳ���������࣬�����Ѿ�ɾ���ˡ�"):exit sub
		end if
		yxbbs.execute("update [yx_board] set boardname='"&boardname&"' where boardid="&boardid)
		call suc("","�ɹ��İ���̳�������Ƹ�Ϊ <b>"&boardname&"</b> !","?")
	end if
end sub

sub delclass
	if  yxbbs.execute("select count(boardid) from[yx_board] where parentid="&yxbbs.boardid)(0)=0 then
		yxbbs.execute("delete from[yx_board] where parentid=0 and boardid="&yxbbs.boardid)
		call suc("","�ɹ���ɾ���˷���!","?")
	end if
end sub

sub addboard
	if yxbbs.execute("select boardid from [yx_board] where depth=0").eof then
		call goback("","û�з��಻��������̳������<a href=board.asp?action=addclass> �����ӷ���</a>")
		exit sub
	end if%>
        <form method=post style="margin:0" action="?action=saveboard">
        <div class="ta">
        <div class="th jz">��̳����</div>
        <div class="td1 h40">��̳���ƣ�</div>
        <div class="td2 h40"><input type=text name="boardname" size=30></div>
        <div class="td1 h40">��־ͼƬ��<br />��̳����logo��ַ�����Բ���</div>
        <div class="td2 h40"><input type=text name="boardimg" size=30></div>
        <div class="td1 h40">��̳������<br />���Ӷ��������@@�ָ����磺name@@name</div>
        <div class="td2 h40"><input type=text name="boardadmin" size=30></div>
        <div class="td1 h70">��̳���ܣ�</div>
        <div class="td2 h70 w446"><textarea rows="3" name="introduce"></textarea></div>
	<div class="td1 h40">���ڷ������̳��</div>
        <div class="td2 h40"><select size="1" name="parentid"><%=yxbbs.boardidlist(yxbbs.boardid,20)%></select></div>
	<div class="td1 h40">�����û��ȼ���</div>
        <div class="td2 h40"><select size="1" name="boardgrade">
	<option value="0" selected>0�����ο�</option>
	<%
	dim ug
	set ug=conn.execute("select gradenum,gradename from yx_usergrade order by gradenum asc")
	while not ug.eof
		response.write "<option value="&ug(0)&">"&ug(0)&"����"&ug(1)&"</option>"
	ug.movenext
	wend
	set ug=nothing
	%>
	</select></div>
        <div class="td1 h40">�Ƿ���֤���棺<br />��֤��ֻ�а�������֤��Ա</div>
        <div class="td2 h40"><input type="radio" value="1" name="boardtype">��&nbsp; <input name="boardtype" type="radio" value="0" checked="true">��</div>
        <div class="td1 h40">�������棺<br /> ����֮�󲻿��Է�������</div>
        <div class="td2 h40"><input type="radio" value="1" name="boardlock">��&nbsp; <input name="boardlock" type="radio" value="0" checked="true">��</div>
<div style="clear: both;"></div><div class="tf jz"><input type="submit" value=" �� �� ">&nbsp;&nbsp;<input type="reset" value=" �� �� "></div>
 </div></form>
	<%
end sub

sub saveboard
	dim boardname,introduce,boardadmin,boardimg,boardgrade,boardlock,boardtype,parentid,newboardid,rootid,depth,child,orders,parentstr,i
	boardname=yxbbs.fun.getstr("boardname")
	introduce=yxbbs.fun.getstr("introduce")
	boardadmin=yxbbs.fun.getstr("boardadmin")
	boardgrade=yxbbs.fun.getstr("boardgrade")
	boardlock=yxbbs.fun.getstr("boardlock")
	boardtype=yxbbs.fun.getstr("boardtype")
	boardimg=yxbbs.fun.getstr("boardimg")
	parentid=yxbbs.fun.getstr("parentid")
	if not isnumeric(parentid) or boardname="" or introduce="" then
		call goback("",""):exit sub
	end if
	set rs=yxbbs.execute("select max(boardid) from [yx_board]")
	if rs.eof or rs.bof then
		call goback("","û�з��಻��������̳������<a href=board.asp?action=addclass> �����ӷ���</a>")
		exit sub
	else
		newboardid=rs(0)+1
	end if
	rs.close
	set rs=yxbbs.execute("select rootid,depth,child,orders,parentstr,parentid from[yx_board] where boardid="&parentid&"")
	if rs.eof or rs.bof then
		call goback("ϵͳ���������","û��ָ���������̳��")
		exit sub
	end if
	rootid=rs(0)
	depth=rs(1)
	child=rs(2)
	orders=rs(3)
	parentstr=rs(4)
	rs.close
	if depth+1>maxdepth then
		call goback("","���ǵ���̳��ʵ�����ã�����̳���������ֻ����" & maxdepth & "����̳!")
		exit sub
	end if
	if parentstr=0 then
		parentstr=parentid
	else
		parentstr=parentstr & "," & parentid
	end if
	yxbbs.execute("insert into [yx_board](boardid,boardname,introduce,boardadmin,boardimg,boardsetting,boardlock,boardtype,rootid,depth,parentid,parentstr,orders,child)values("&newboardid&",'"&boardname&"','"&introduce&"','"&boardadmin&"','"&boardimg&"','0,gif|jpg|swf|zip|rar|bmp,30,30720,10,0,0,0,0,0,0,4',"&boardlock&","&boardtype&","&rootid&","&depth+1&","&parentid&",'"&parentstr&"',"&orders+1&",0)")
	if boardadmin <> "" then call addmaster(boardadmin,"none",0)
	if parentid<>0 then
		if depth>0 then
			'���ϼ�������ȴ���0��ʱ��Ҫ�����丸�ࣨ����ĸ��ࣩ�İ��������������
			for i=1 to depth
				'�����丸�������
				yxbbs.execute("update [yx_board] set child=child+1 where boardid="&parentid)
				'�õ��丸��ĸ���İ���id
				set rs=yxbbs.execute("select parentid from [yx_board] where boardid="&parentid)
				if not (rs.eof and rs.bof) then
					parentid=rs(0)
				end if
				rs.close
				'��ѭ����������1�������е����һ��ѭ����ʱ��ֱ�ӽ��и���
				if i=depth then
				yxbbs.execute("update [yx_board] set child=child+1 where boardid="&parentid)
				end if
			next
			'���¸ð��������Լ����ڱ���Ҫ��ͬ�ڱ������µİ����������
			yxbbs.execute("update [yx_board] set orders=orders+1 where rootid="&rootid&" and orders>"&orders)
			yxbbs.execute("update [yx_board] set orders="&orders&"+1 where boardid="&newboardid&"")
		else
			'���ϼ��������Ϊ0��ʱ��ֻҪ�����ϼ����������
			yxbbs.execute("update [yx_board] set child=child+1 where boardid="&parentid)
			set rs=yxbbs.execute("select max(orders) from [yx_board]")
			yxbbs.execute("update [yx_board] set orders="&rs(0)&"+1 where boardid="&newboardid )
			rs.close
		end if
	end if
	call suc("","�ɹ�����������̳ <b>"&boardname&"</b> !","board.asp")	
end sub

sub editboard
	dim boardname,boardlock,boardtype,introduce,boardimg,parentid,boardadmin,boardgrade
	set rs=yxbbs.execute("select parentid,boardname,boardadmin,boardgrade,boardlock,boardtype,introduce,boardimg from[yx_board] where boardid="&yxbbs.boardid&"")
	if rs.eof then
		call goback("","�ð��治���ڣ������Ѿ�ɾ����")
		exit sub
	else
		parentid=rs(0)
		boardname=rs(1)
		boardadmin=rs(2)
		boardgrade=rs(3)
		boardlock=rs(4)
		boardtype=rs(5)
		introduce=rs(6)
		boardimg=rs(7)
	end if
	rs.close%><form method=post style="margin:0" action="?action=saveeditboard">
         <div class="ta">
        <div class="th jz">��̳�޸�</div>
        <div class="td1 h40">��̳���ƣ�</div>
        <div class="td2 h40"><input name="boardid" value="<%=yxbbs.boardid%>" type="hidden"><input name="boardname" type=text value="<%=boardname%>" size=30></div>
        <div class="td1 h40">��־ͼƬ��<br />��̳����logo��ַ�����Բ���</div>
        <div class="td2 h40"><input type=text name="boardimg" value="<%=boardimg%>" size=30></div>
        <div class="td1 h40">��̳������<br />���Ӷ��������@@�ָ����磺name@@name</div>
        <div class="td2 h40"><input type=text name="boardadmin" size=30 value="<%=boardadmin%>"><input type="hidden" name="oldboardadmin" size=30 value="<%=boardadmin%>"></div>
        <div class="td1 h70">��̳���ܣ�</div>
        <div class="td2 h70 w446"><textarea rows="3" name="introduce"><%=introduce%></textarea></div>
	<div class="td1 h40">���ڷ������̳��</div>
        <div class="td2 h40"><select size="1" name="parentid"><%=yxbbs.boardidlist(parentid,20)%></select></div>
	<div class="td1 h40">�����û��ȼ���</div>
        <div class="td2 h40"><select size="1" name="boardgrade">
	<option value="0" selected>0�����ο�</option>
	<%
	dim ug
	set ug=conn.execute("select gradenum,gradename from yx_usergrade order by gradenum asc")
	while not ug.eof
		if ug(0)=boardgrade then
			response.write "<option value="&ug(0)&" selected>"&ug(0)&"����"&ug(1)&"</option>"
		else
			response.write "<option value="&ug(0)&">"&ug(0)&"����"&ug(1)&"</option>"
		end if
	ug.movenext
	wend
	set ug=nothing
	%></select></div>
        
        <div class="td1 h40">�Ƿ���֤���棺<br />��֤��ֻ�а�������֤��Ա</div>
        <div class="td2 h40"><input type="radio" value="1" name="boardtype" <%if boardtype=1 then response.write "checked=""true"""%>>��&nbsp; <input type="radio" value="0" name="boardtype" <%if boardtype=0 then response.write "checked=""true"""%>>��</div>

<div class="td1 h40">�������棺<br /> ����֮�󲻿��Է�������</div>
        <div class="td2 h40"><input type="radio" value="1" name="boardlock" <%if boardlock=1 then response.write "checked=""true""" %>>��&nbsp; <input type="radio" value="0" name="boardlock" <%if boardlock=0 then response.write "checked=""true"""%>>��</div>
<div style="clear: both;"></div><div class="tf jz"><input type="submit" value=" �� �� ">&nbsp;&nbsp;<input type="reset" value=" �� �� "></div>
</div></form>
	<%
end sub

sub saveeditboard
	dim boardid,boardname,introduce,boardadmin,oldboardadmin,boardgrade,boardlock,boardtype,boardimg,parentid,rootid,depth,child,orders,parentstr,i
	dim newparentid,boardnum,p_rs
	boardid=yxbbs.fun.getstr("boardid")
	boardname=yxbbs.fun.getstr("boardname")
	introduce=yxbbs.fun.getstr("introduce")
	boardadmin=yxbbs.fun.getstr("boardadmin")
	boardgrade=yxbbs.fun.getstr("boardgrade")
	oldboardadmin=yxbbs.fun.getstr("oldboardadmin")
	boardlock=yxbbs.fun.getstr("boardlock")
	boardtype=yxbbs.fun.getstr("boardtype")
	boardimg=yxbbs.fun.getstr("boardimg")
	newparentid=yxbbs.fun.getstr("parentid")
	if not isnumeric(newparentid) or boardname="" or introduce="" then
		call goback("",""):exit sub
	elseif boardid=newparentid then
		call goback("","������̳����ָ���Լ���"):exit sub
	end if
	set rs=yxbbs.execute("select rootid,depth,child,orders,parentid,parentstr from[yx_board] where boardid="&boardid)
	if rs.eof or rs.bof then
		call goback("ϵͳ������","�ð��治���ڣ������Ѿ�ɾ���ˣ�")
		exit sub
	end if
	rootid=rs(0)
	depth=rs(1)
	child=rs(2)
	orders=rs(3)
	parentid=rs(4)
	parentstr=rs(5)
	rs.close
	if parentid=0 then
		call goback("ϵͳ������","���಻������"):exit sub
	elseif int(newparentid)<>int(parentid) then
		'�ж���ָ������̳�Ƿ���������̳
		set rs=yxbbs.execute("select boardid from [yx_board] where parentstr like '%"&parentstr&","&boardid&"%' and boardid="&newparentid)
		if not (rs.eof and rs.bof) then
			call goback("","������ָ���ð������������̳��Ϊ������̳")
			exit sub
		end if
		rs.close
		'�����ѡ�ĸ���
		set p_rs=yxbbs.execute("select * from [yx_board] where boardid="&newparentid)
			if p_rs("depth")+1> maxdepth or (child>0 and p_rs("depth")+2>maxdepth) then
			p_rs.close:set p_rs=nothing
			call goback("","����̳���������ֻ����" & maxdepth & "����̳.")
			exit sub
		end if
	end if
	yxbbs.execute("update [yx_board] set boardname='"&boardname&"',introduce='"&introduce&"',boardadmin='"&boardadmin&"',boardgrade="&boardgrade&",boardlock="&boardlock&",boardtype="&boardtype&",boardimg='"&boardimg&"' where boardid="&boardid&"")
	if boardadmin <> oldboardadmin then call addmaster(boardadmin,oldboardadmin,1)
	if int(newparentid)<>int(parentid) then
	'��һ������̳�ƶ�����������̳��
	'�����ָ������̳�������Ϣ
	'�õ�������������
	parentstr=parentstr & ","
	boardnum=yxbbs.execute("select count(*) from [yx_board] where parentstr like '%"&parentstr & boardid&"%'")(0)
	if isnull(boardnum) then boardnum=1
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	yxbbs.execute("update [yx_board] set orders=orders + "&boardnum&"+1  where rootid="&p_rs("rootid")&" and orders>"&p_rs("orders")&"")
	'���µ�ǰ��������
	if p_rs("parentstr")="0" then
		yxbbs.execute("update [yx_board] set depth="&p_rs("depth")&"+1,orders="&p_rs("orders")&"+1,rootid="&p_rs("rootid")&",parentid="&newparentid&",parentstr='" & p_rs("boardid") & "' where boardid="&boardid)
	else
		yxbbs.execute("update [yx_board] set depth="&p_rs("depth")&"+1,orders="&p_rs("orders")&"+1,rootid="&p_rs("rootid")&",parentid="&newparentid&",parentstr='" & p_rs("parentstr") & ","& p_rs("boardid") &"' where boardid="&boardid)
	end if
	dim tempparentstr
	i=1
	'����������ͬʱ����ƶ�����i
	'����������������������
	'���Ϊԭ����ȼ��ϵ�ǰ������̳�����
	set rs=yxbbs.execute("select * from [yx_board] where parentstr like '%"&parentstr & boardid&"%' order by orders")
	do while not rs.eof
	i=i+1
	if p_rs("parentstr")="0" then'����丸��Ϊ�࣬��ô�������İ�������
		tempparentstr=p_rs("boardid") & "," & replace(rs("parentstr"),parentstr,"")
	else
		tempparentstr=p_rs("parentstr") & "," & p_rs("boardid") & "," & replace(rs("parentstr"),parentstr,"")
	end if
	yxbbs.execute("update [yx_board] set depth=depth+"&p_rs("depth")&"-"&depth&"+1,orders="&p_rs("orders")&"+"&i&",rootid="&p_rs("rootid")&",parentstr='"&tempparentstr&"' where boardid="&rs("boardid"))
	rs.movenext
	loop
	rs.close
	dim tempparentid,ii
	tempparentid=newparentid
	if rootid=p_rs("rootid") then'��ͬһ�������ƶ�
		'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
		'�����丸�������
		yxbbs.execute("update [yx_board] set child=child+"&i&" where (not parentid=0) and boardid="&tempparentid)
		for ii=1 to p_rs("depth")
			'�õ��丸��ĸ���İ���id
			set rs=yxbbs.execute("select parentid from [yx_board] where (not parentid=0) and boardid="&tempparentid)
			if not (rs.eof and rs.bof) then
				tempparentid=rs(0)
				'�����丸��ĸ��������
			yxbbs.execute("update [yx_board] set child=child+"&i&" where (not parentid=0) and boardid="&tempparentid)
			''''yxbbs.execute("update [yx_board] set child=child+"&i&" where (not parentid=0) and boardid in������)
			end if
		next
		'������ԭ���������
		yxbbs.execute("update [yx_board] set child=child-"&i&" where (not parentid=0) and boardid="&parentid)
		'������ԭ��������̳����
		for ii=1 to depth
			'�õ���ԭ����ĸ���İ���id
			set rs=yxbbs.execute("select parentid from [yx_board] where (not parentid=0) and boardid="&parentid)
			if not (rs.eof and rs.bof) then
				parentid=rs(0)
				'������ԭ����ĸ��������
				yxbbs.execute("update [yx_board] set child=child-"&i&" where (not parentid=0) and  boardid="&parentid)
			end if
		next
	else
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
		yxbbs.execute("update [yx_board] set child=child+"&i&" where boardid="&tempparentid)
		for ii=1 to p_rs("depth")
			'�õ��丸��ĸ���İ���id
			set rs=yxbbs.execute("select parentid from [yx_board] where boardid="&tempparentid)
			if not (rs.eof and rs.bof) then
				tempparentid=rs(0)
				'�����丸��ĸ��������
			yxbbs.execute("update [yx_board] set child=child+"&i&" where  boardid="&tempparentid)
			end if
		next
	'������ԭ���������
	yxbbs.execute("update [yx_board] set child=child-"&i&" where boardid="&parentid)
	'������ԭ�����������������
	yxbbs.execute("update [yx_board] set orders=orders-"&i&" where rootid="&rootid&" and orders>"&orders)
	'������ԭ��������̳����
		for ii=1 to depth
			'�õ���ԭ����ĸ���İ���id
			set rs=yxbbs.execute("select parentid from [yx_board] where boardid="&parentid)
			if not (rs.eof and rs.bof) then
				parentid=rs(0)
				'������ԭ����ĸ��������
				yxbbs.execute("update [yx_board] set child=child-"&i&" where boardid="&parentid)
			end if
		next
	end if
	p_rs.close:set p_rs=nothing
  end if
	call suc("","��̳�޸ĳɹ� !","board.asp")	
end sub

sub delboard
	dim alltable,i,ii,depth,parentid,rootid,orders
	set rs=yxbbs.execute("select depth,parentid,rootid,orders,child from[yx_board] where boardid="&yxbbs.boardid)
	if rs.eof then 
		call goback("ϵͳ����","�����ڣ���̳�����Ѿ�ɾ���� !")
		exit sub
	elseif rs(4)>0 then
		call goback("ϵͳ����","����̳����������̳������ɾ�� !")
		exit sub
	else
		depth=rs(0)
		parentid=rs(1)
		rootid=rs(2)
		orders=rs(3)
	end if
	rs.close
	alltable=split(yxbbs.bbstable(0),",")
	for i=0 to ubound(alltable)
		yxbbs.execute("delete from [yx_bbs"&alltable(i)&"] where boardid="&yxbbs.boardid&"")
	next

	set rs=yxbbs.execute("select topicid from[yx_topic] where boardid="&yxbbs.boardid&" and isvote ")
	do while not rs.eof
		yxbbs.execute("delete from [yx_topicvote] where topicid="&rs(0))
		yxbbs.execute("delete from [yx_topicvoteuser] where topicid="&rs(0))
	rs.movenext
	loop
	rs.close
	'ɾ�������¼
	yxbbs.execute("delete from[yx_topic] where boardid="&yxbbs.boardid)
	yxbbs.execute("delete from[yx_board] where boardid="&yxbbs.boardid)
	'�����丸��İ�����
	yxbbs.execute("update [yx_board] set child=child-1 where boardid="&parentid)
	'�����丸���������������
	yxbbs.execute("update [yx_board] set orders=orders-1 where rootid="&rootid&" and orders>"&orders)
	'������������̳����
		for ii=1 to depth
			'�õ��丸��ĸ���İ���id
			set rs=yxbbs.execute("select parentid from [yx_board] where boardid="&parentid)
			if not (rs.eof and rs.bof) then
				parentid=rs(0)
				'�����丸��ĸ��������
				yxbbs.execute("update [yx_board] set child=child-1 where boardid="&parentid)
			end if
			rs.close
		next
	call suc("","�ɹ���ɾ����̳���� (��������̳����������)!","board.asp")	
end sub

sub classorders
	dim boardid
	set rs=yxbbs.execute("select boardid,boardname,rootid from[yx_board] where depth=0 order by rootid")
	if rs.eof then
		call goback("","��̳û�з��࣡����<a href=?action=addclass> ���ӷ���</a>")
		exit sub
	end if
	%><form method=post style="margin:0" action="?action=saveclassorders">
        <div class="ta">
	<div class="th jz">��������</div>
       
        <%
	do while not rs.eof
	response.write"<div class=""td1 h30"">&nbsp;"&rs(1)&"</div><div class=""td2 h30""><input name='boardid' type='hidden' value='"&rs(0)&"'><input name='rootid' type='hidden' value='"&rs(2)&"'><input type=text name='newrootid' value='"&rs(2)&"' size=4 ></div>"
	rs.movenext
	loop
	%><div style="clear: both;"></div><div class="tf jz"><input type="submit" value=" �� �� ">&nbsp;&nbsp;<input type="reset" value=" �� �� "></div>
        </div></form>
	
	<%rs.close
end sub

sub saveclassorders
	dim boardid,rootid,newrootid,temp,i
	temp=","
	for i=1 to request.form("boardid").count
		boardid = request.form("boardid")(i)
		rootid = request.form("rootid")(i)
		newrootid = request.form("newrootid")(i)
		if instr(temp,","&newrootid&",")>0 then 
			call goback("�������","��������������ֲ���һ��!")
			exit sub
		end if
		temp=temp&newrootid&","
		if not isnumeric(boardid) or not isnumeric(newrootid) then
			call goback("�������","����������д!")
			exit sub
		end if
	next
	for i=1 to request.form("boardid").count
		boardid = request.form("boardid")(i)
		rootid = request.form("rootid")(i)
		newrootid = request.form("newrootid")(i)
		if rootid<>newrootid then
			yxbbs.execute("update [yx_board]set rootid="&newrootid&" where boardid="&boardid)
			temp=boardid
			yxbbs.execute("update [yx_board] set rootid="&newrootid&" where parentstr like '%"&temp&"%' and rootid="&rootid&"")
		end if
	next
	call suc("","����ɹ���","?")	
end sub
sub addmaster(s,o,n)
	dim arr,oarr,i,sql
	dim gradenum,gradename,gradepic
	set rs=conn.execute("select gradenum,gradename,gradepic from yx_usergrade where classid = 3 order by userminpostnum desc")
	if not (rs.eof and rs.bof) then
		gradenum = rs(0)
		gradename = rs(1)
		gradepic = rs(2)
	end if
	randomize
	arr = split(s,"@@")
	oarr = split(o,"@@")
	set rs = server.createobject("adodb.recordset")
	for i = 0 to ubound(arr)
		sql = "select * from [yx_user] where name = '" & arr(i) & "'"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			call goback("",arr(i) &"����û���û��ע��!")
		else
			if rs("classid") = 5 then
				rs("gradenum") = gradenum
				rs("classid") = 3
				rs.update
			end if
		end if
		rs.close
	next
	'�ж�ԭ���������������Ƿ񻹵��ΰ�������û�е����򳷻����û�ְλ��
	if n = 1 then
		dim iboardmaster
		dim usergrade, article
		iboardmaster = false
		for i = 0 to ubound(oarr)
			set rs = yxbbs.execute("select boardadmin from yx_board")
			do while not rs.eof
				if instr("|" & trim(rs(0)) & "|","|" & trim(oarr(i)) & "|") > 0 then
					iboardmaster = true
					exit do
				end if
				rs.movenext
			loop
			if not iboardmaster then
				set rs = yxbbs.execute("select id,classid,essaynum from [yx_user] where name='" & trim(oarr(i)) & "' or classid=3")
				if not (rs.eof and rs.bof) then
					'if rs(1) > 2 then
						if not isnumeric(rs(2)) or rs(2) = "" then
							article = 0
						else
							article = cstr(rs(2))
						end if
						'ȡ��Ӧע���Ա�ĵȼ�
						set usergrade = yxbbs.execute("select top 1 gradenum,gradename,gradepic from yx_usergrade where userminpostnum <= " & article & " and not userminpostnum = -1 and classid = 5 order by userminpostnum desc")
						if not (usergrade.eof and usergrade.bof) then
							yxbbs.execute("update [yx_user] set gradenum="&usergrade(0)&",classid=5 where id="&rs(0))
						end if
						usergrade.close:set usergrade = nothing
					'end if
				end if
			end if
			iboardmaster = false
		next
	end if
	set rs = nothing
end sub

sub boardunite
	%><form method=post style="margin:0" action="?action=saveboardunite">
        <div class="ta">
        <div class="th jz">��̳�ϲ�</div>
        <div class="td w772">����̳�� 
	  <select size="1" name="boardid"><option value='0'>��ѡ��ԭ��̳</option><%=yxbbs.boardidlist(0,0)%></select> �ϲ�����̳�� <select size="1" name="newboardid"><option value='0'>��ѡ��Ŀ����̳</option><%=yxbbs.boardidlist(0,0)%></select> �� <input type="submit" onclick=checkclick('�����󽫲��ָܻ�����ȷ��Ҫ�ϲ���') value=" ��̳�ϲ� ">
	<br /><br /><b>ע�����</b><font color=red>�˲������ɻָ��������ز�����</font><br />
	���಻�ܲ��������ܺ������µ���̳�ϲ���<br />
	�ϲ���ԭ��̳(����������̳)����ɾ������������(����������̳������)��ת�Ƶ�ָ����Ŀ����̳��</div>
	<div style="clear: both;"></div></div></form>
	<%
end sub

sub saveboardunite
	dim boardid,newboardid,tempparentstr,tempparentid,rs1
	dim i,alltable
	dim parentstr,depth,parentid,child,rootid
	boardid=yxbbs.fun.getstr("boardid")
	newboardid=yxbbs.fun.getstr("newboardid")
	if boardid="" or newboardid="" then
		call goback("","����ָ����̳���ٽ��кϲ���")
		exit sub
	elseif boardid=newboardid then
		call goback("","ͬһ����̳���úϲ��ˣ�")
		exit sub
	end if

	set rs=yxbbs.execute("select parentstr,boardid,depth,parentid,child,rootid from yx_board where boardid="&boardid)
	if rs(2)="0" then
		call goback("ϵͳ����","���಻�����ϲ�������")
		exit sub
	end if
	parentstr=rs(0) & "," & rs(1)
	parentid=rs(3)
	tempparentstr=rs(1)
	depth=rs(2)
	child=rs(4)+1
	rootid=rs(5)
	rs.close
	tempparentid=parentid
	'�ж��Ƿ�ϲ���������̳
	set rs=yxbbs.execute("select boardid from [yx_board] where boardid="&newboardid&" and parentstr like '%"&parentstr&"%'")
	if not (rs.eof and rs.bof) then
		call goback("","���ܽ���̳�ϲ�����������̳��!")
		exit sub
	end if
	rs.close
	'�õ�ȫ��������̳id
	i=0
	set rs=yxbbs.execute("select boardid from [yx_board] where rootid="&rootid&" and parentstr like '%"&parentstr&"%'")
	do while not rs.eof
		if i=0 then
			tempparentstr=rs(0)
		else
			tempparentstr=tempparentstr & "," & rs(0)
		end if
		i=i+1
		rs.movenext
	loop
	if i>0 then
		parentstr=tempparentstr & "," & boardid
	else
		parentstr=boardid
	end if
	'������ԭ��������̳������
	yxbbs.execute("update [yx_board] set child=child-"&child&" where boardid="&tempparentid)
	'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
	for i=1 to depth
		'�õ��丸��ĸ���İ���id
		set rs=yxbbs.execute("select parentid from [yx_board] where boardid="&tempparentid)
		if not (rs.eof and rs.bof) then
			tempparentid=rs(0)
			yxbbs.execute("update [yx_board] set child=child-"&child&" where boardid="&tempparentid)
		end if
	next
	'������̳��������
	alltable=split(yxbbs.bbstable(0),",")
	for i=0 to ubound(alltable)
		yxbbs.execute("update [yx_bbs"&alltable(i)&"] set boardid="&newboardid&" where boardid in ("&parentstr&")")
	next
	yxbbs.execute("update [yx_topic] set boardid="&newboardid&" where boardid in ("&parentstr&")")
	'ɾ�����ϲ���̳
	set rs=yxbbs.execute("select sum(essaynum),sum(topicnum),sum(todaynum) from [yx_board] where rootid="&rootid&" and boardid in ("&parentstr&")")
	yxbbs.execute("delete from [yx_board] where rootid="&rootid&" and boardid in ("&parentstr&")")
	'��������̳���Ӽ���
	yxbbs.execute("update [yx_board] set essaynum=essaynum+"&rs(0)&",topicnum=topicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid ="&newboardid&"")
	'�����ϼ����
	set rs1=yxbbs.execute("select depth,parentstr,boardid from [yx_board] where boardid="&newboardid)
	if rs1(0)>1 then
	parentstr=rs1(0)
	yxbbs.execute("update [yx_board] set essaynum=essaynum+"&rs(0)&",topicnum=topicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid in ("&parentstr&")")
	end if
	rs1.close:set rs1=nothing
	rs.close
	call suc("","�ϲ��ɹ����Ѿ���ԭ��̳���������£����������Ӻϲ���Ŀ����̳��","?")
end sub
sub ordersboard
	dim boardid,parentid,rootid,orders,parentstr,i,boardnum,p_rs,boardname
	boardid=yxbbs.boardid
	set rs=yxbbs.execute("select orders,parentid,parentstr,boardname from[yx_board] where boardid="&yxbbs.boardid)
	if rs.eof or rs.bof then
		call goback("ϵͳ������","�ð��治���ڣ������Ѿ�ɾ���ˣ�")
		exit sub
	end if
	orders=rs(0)
	parentid=rs(1)
	parentstr=rs(2)
        
	rs.close
	'������Ϊ��ʱ
	if parentid=0 then call goback("ϵͳ������","����id������"):exit sub
	'�õ�������������
	parentstr=parentstr & ","
	boardnum=yxbbs.execute("select count(*) from [yx_board] where parentstr like '%"&parentstr & boardid&"%'")(0)
	if isnull(boardnum) then boardnum=1
	'��ø�����Ϣ
	set p_rs=yxbbs.execute("select * from [yx_board] where boardid="&parentid)
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	yxbbs.execute("update [yx_board] set orders=orders + "&boardnum&"+1  where rootid="&p_rs("rootid")&" and orders>"&p_rs("orders")&"")
	'���µ�ǰ��������
	yxbbs.execute("update [yx_board] set orders="&p_rs("orders")&"+1 where boardid="&boardid)
	dim tempparentstr
	i=1
	'����������ͬʱ����ƶ�����i
	'����������������������
	set rs=yxbbs.execute("select * from [yx_board] where parentstr like '%"&parentstr & boardid&"%' order by orders")
	do while not rs.eof
	i=i+1
	if p_rs("parentstr")="0" then'����丸��Ϊ�࣬��ô�������İ�������
		tempparentstr=p_rs("boardid") & "," & replace(rs("parentstr"),parentstr,"")
	else
		tempparentstr=p_rs("parentstr") & "," & p_rs("boardid") & "," & replace(rs("parentstr"),parentstr,"")
	end if
	yxbbs.execute("update [yx_board] set orders="&p_rs("orders")&"+"&i&",parentstr='"&tempparentstr&"' where boardid="&rs("boardid"))
	rs.movenext
	loop
	rs.close
	p_rs.close
	set p_rs=nothing
	call suc("","�ɹ��Ľ��ð�����ǰ!","?")
	response.end
end sub

sub passuser
	set rs=yxbbs.execute("select passuser,boardname from [yx_board] where boardid="&yxbbs.boardid&" and parentid<>0")
	if rs.eof then
		call goback("","����̳�����Ͳ�����֤��̳������������֤�û���")
		exit sub
	end if%>
	<form method=post style='margin:0' action="?action=savepassuser">
	<div class="ta">
        <div class="th jz">�޸���̳��֤�û�</div>
        <div class="td1">������̳��</div>
        <div class="td2"><input name="boardid" value="<%=yxbbs.boardid%>" type="hidden"><%=rs("boardname")%></div>
        <div class="td1 h50"><b>ͨ����֤���û���</b><br />���û�֮���á�|������<br />�벻Ҫʹ�ûس���enter</div>
        <div class="td2 h50 w446"><textarea name="passuser" rows="2"><%=rs("passuser")%></textarea></div>
	<div style="clear: both;"></div><div class="tf"><input type="submit" value=" �� �� "> &nbsp;&nbsp;<input type="reset" name="submit" value=" �� �� "></div></div></form>
	<%rs.close
end sub

sub savepassuser
	dim passusers,boardid
	boardid=yxbbs.fun.getstr("boardid")
	passusers=trim(replace(request.form("passuser"),"'",""))
	passusers=replace(passusers,chr(10), "")
	passusers=replace(passusers,chr(13), "")
	yxbbs.execute("update [yx_board] set passuser='"&passusers&"' where boardid="&boardid&" and parentid<>0")
	cache.name="boardinfo"
	cache.clean()
	call suc("","�ɹ��ĸ����˸���̳����֤��Ա��","?")
end sub
%>