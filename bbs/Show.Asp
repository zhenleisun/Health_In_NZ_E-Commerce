<!--#include file="Inc/SysConfig.Asp"-->
<!--#include file="Inc/ubb_Cls.Asp"-->
<!--#include file="Inc/page_Cls.Asp"-->
<script type="text/javascript">
var y=0;function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){y++;if (y>1) {alert('�������ڷ����������ĵȴ���');return false;}this.document.yimxu.submit();this.document.yimxu.BtnPost.disabled=true;}}
</script>
<%
Dim TopicCaption,TopicUserName,TopicTopType,TopicIsGood,TopicIsLock,TopicIsVote,TopicHits,TopicReplyNum
Dim YxBBs_ubb,Page,PageString,PageInfo,ID,UserName,BBSID,GCode
ID=YxBBs.CheckNum(Request.QueryString("ID"))
Call GetTopicInfo()
Call ShowListTop()
If TopicIsVote Then ShowVote()
Set YxBBs_ubb=New YxBBsubb_Cls
Call ShowBBS()
Set YxBBs_ubb=Nothing
ShowListBottom()
FastReply()
YxBBs.Footer()
YxBBs.Execute("UPDATE [YX_Topic] SET Hits=Hits+1 WHERE TopicID="&ID&"")
Set YxBBs=Nothing
	
Sub GetTopicInfo()
	Dim Rs,Arr_Rs
	Set Rs=YxBBs.Execute("Select TopicID,Caption,Name,TopType,IsGood,IsLock,isVote,Hits,ReplyNum,AddTime,SqlTableID,BoardID From [YX_Topic] where TopicID="&ID&" and IsDel=False")
	IF Rs.eof then YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Arr_Rs=Rs.GetRows(1)
	Rs.Close
	Set Rs=Nothing
	TopicCaption=YxBBs.Fun.HtmlCode(Arr_Rs(1,0))
	TopicUserName=Arr_Rs(2,0)
	TopicTopType=Arr_Rs(3,0)
	TopicIsGood=Arr_Rs(4,0)
	TopicIsLock=Arr_Rs(5,0)
	TopicIsVote=Arr_Rs(6,0)
	TopicHits=Arr_Rs(7,0)
	TopicReplyNum=Arr_Rs(8,0)
	YxBBs.TB=Arr_Rs(10,0)
	YxBBs.BoardID=Arr_Rs(11,0)
	YxBBs.CheckBoard()
	YxBBs.Head(Arr_Rs(1,0))
	If YxBBs.MyName<>TopicUserName Then
		If Cint(YxBBs.ClassSetting(1))=0 Then YxBBs.Error("��û�в鿴�������ӵ�Ȩ�ޣ�")
		If TopicIsGood Then If Cint(YxBBs.ClassSetting(2))=0 Then YxBBs.Error("��û������������ӵ�Ȩ�ޣ�")
	End If
End Sub

Function TopicMood()
	Dim Temp,TopicImg
	TopicImg=Split(YxBBs.Template.ReadTemplate("����״̬ͼƬ"),"|")
	IF TopicIsGood Then Temp="<td>"&TopicImg(3)&" <font color=red>��������</font></td>"
	IF TopicIsVote then Temp=Temp&"<td>"&TopicImg(4)&" <font color=red>ͶƱ����</font></td>"
	IF TopicIsLock then Temp=Temp&"<td>"&TopicImg(7)&" <font color=red>��������</font></td>"
	IF TopicTopType=3 then Temp=Temp&"<td>"&TopicImg(2)&" <font color=red>�ö�����</font></td>"
	IF TopicTopType=4 then Temp=Temp&"<td>"&TopicImg(1)&" <font color=red>���ö�����</font></td>"
	IF TopicTopType=5 then Temp=Temp&"<td>"&TopicImg(0)&" <font color=red>���ö�����</font></td>"
	IF Temp<>"" Then Temp="<Div  align='right'><table border=1 style='border-collapse: collapse' cellpadding='5' cellspacing='0' bordercolor=#DEDEDE bgcolor=ffffff><tr>"&Temp&"</tr></table></div>"
	TopicMood=Temp
End Function

Function ShowListTop()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("���Ӷ���")
	Temp=Replace(Temp,"{������ť}",SayBar())
	Temp=Replace(Temp,"{�����}",TopicHits+1)
	Temp=Replace(Temp,"{����}",TopicCaption)
	Response.Write(Temp)
End Function

Function ShowListBottom()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("���ӵײ�")
	Temp=Replace(Temp,"{��ҳ}",PageInfo)
	Response.Write(Temp)
End Function

Function SayBar()
	Dim Temp,PostPic
	PostPic=Split(YxBBs.Template.ReadTemplate("����ͼƬ"),"||")
	Temp="<a href='Say.Asp?BoardID="&YxBBs.BoardID&"'>"&PostPic(0)&"</a> <a href='Say.Asp?action=vote&BoardID="&YxBBs.BoardID&"'>"&PostPic(1)&"</a>"
	If Not TopicIsLock or YxBBs.ClassID=1 Then
		Temp=Temp&" <a href='Say.Asp?Action=Reply&BoardID="&YxBBs.BoardID&"&ID="&ID&"'>"&PostPic(2)&"</a>"
	End If
	SayBar=Temp
End Function

Function ShowVote()
	Dim Temp,Rs,Arr_Rs,Vote,VoteNum,AllvoteNum,VotePicW,Opt,ClueTxt,CanVote,ii,VoteShow,i
	Set Rs=YxBBs.Execute("Select TopicID,Vote,VoteNum,VoteType,OutTime From [YX_TopicVote] where TopicID="&ID&"")
	If Rs.Eof then Exit Function
	Arr_Rs=Rs.GetRows(1)
	Rs.Close:Set Rs=Nothing
	Vote=Split(Arr_Rs(1,0),"|")
	VoteNum=split(Arr_Rs(2,0),"|")
	CanVote=True
	If Not YxBBs.FoundUser Then
		ClueTxt="�㻹û�е�½�����ܽ���ͶƱ��"
		CanVote=False
	Else
		IF not YxBBs.Execute("Select User From [YX_TopicVoteUser] where User='"&YxBBs.MyName&"' and TopicID="&ID&"").eof then
			ClueTxt="���Ѿ�Ͷ��Ʊ�ˣ�������ͶƱ�ˡ�[<a style='cursor:hand;color:red;' onclick=javascript:window.open('See.Asp?Action=Vote&Id="&Arr_Rs(0,0)&"','open','width=550,height=400,resizable=1,scrollbars=yes,menubar=no,status=yes')>����</a>]"
			CanVote=False
		End If
	End If
	For i=1 to ubound(Vote)
		AllvoteNum=Int(AllvoteNum+VoteNum(i))
	Next
	IF AllVoteNum=0 then AllvoteNum=1
	For i=1 To ubound(Vote)
		ii=ii+1
		VotePicW=VoteNum(i)/AllvoteNum*85
		If CanVote Then
			IF Int(Arr_Rs(3,0))=1 then Opt="<input type=radio value='"&i&"' name='opt'>" Else Opt="<input type=checkbox name='opt"&i&"' value='"&i&"'>"
		End If
		dim Svote
		IF not YxBBs.Execute("Select User From [YX_TopicVoteUser] where User='"&YxBBs.MyName&"' and TopicID="&ID&"").eof then
		Svote="<img border=0 height=8 width='"&VotePicW&"%' src='Images/hr.gif'> <b>"&VoteNum(i)&"</b> Ʊ"
		else
                Svote="<img border=0 height=8 width='50%' src='Images/hr.gif'>  ͶƱ��ɲ鿴���飡"
		End If
		VoteShow=VoteShow&"<tr><td width='5%' height=25 >&nbsp;"&i&"</td><td width='40%'>&nbsp;"&Opt&""&YxBBs.Fun.HtmlCode(Vote(i))&"</td><td width='35%' >&nbsp;"&Svote&"</td></tr>"

	Next
	If DateDiff("s",YxBBs.NowBbsTime,Arr_Rs(4,0))<0 then
		ClueTxt="��ͶƱ�Ѿ����ڣ����ܽ���ͶƱ��"
		CanVote=False
	End If
	IF CanVote then
		ClueTxt="<input type='submit' value='Ͷ Ʊ (ͶƱ���ܿ��꾡���)' ID='BtnPost'>"
	End IF
	ClueTxt=ClueTxt&" [ ��ֹʱ�䣺"&Arr_Rs(4,0)&" ]"
	Temp=YxBBs.Template.ReadTemplate("��ʾͶƱ")
	Temp=Replace(Temp,"{ͶƱѡ��}","<form style='margin:0' method=POST onsubmit='BtnPost.disabled = true;' action=VoteSubmit.Asp?Action=Vote&id="&ID&"&type="&Arr_Rs(3,0)&">"&VoteShow)
	Temp=Replace(Temp,"{ͶƱ��Ϣ}",ClueTxt&"</form>")
	Response.Write(Temp)
End Function

Function ShowBBS()
	Dim arr_Rs,i,pages,rs,GradeName,Gradepic
	Dim Template,TempUserL,Temp,TempStr,Lou,Sign,ShowCaption,ShowMood,Arr_Ip
	Page = Request.QueryString("page")
	Set pages = New Cls_PageView
	pages.strTableName = "[YX_Bbs"&YxBBs.TB&"] As B inner join [YX_User] As U on B.Name=U.Name"
	pages.strPageUrl = "?BoardID="&YxBBs.BoardID&"&ID="&ID&"&TB="&YxBBs.TB
	pages.strFieldsList = "B.BbsID,B.TopicID,B.Caption,B.Content,B.Name,B.AddTime,B.BoardID,B.UbbString,B.IP,U.Id,U.Name,U.IsQQpic,U.QQ,U.Pic,U.Picw,U.Pich,U.GradeNum,U.EssayNum,U.Mail,U.Home,U.Sex,U.Mark,U.Coin,U.Sign,U.Regtime,U.IsShow,U.IsSign,U.RegIp,U.LoginNum,U.Honor,U.TimeSum"
	pages.strCondiction = "B.isDel=False and (B.TopicID="&ID&" or B.ReplyTopicID="&ID&")"
	pages.strOrderList = "B.BbsID"
	pages.strPrimaryKey = "BbsID"
	Pages.CountSQL=TopicReplyNum+1
	'Pages.CountSQL=YxBBs.Execute("SELECT Count(ID) FROM [YX_Bbs"&YxBBs.TB&"] where TopicID="&ID&" or ReplyTopicID="&ID&"")(0)
	pages.intPageSize = 10
	pages.intPageNow = Page
	pages.strCookiesName = "Show_"&ID
	pages.Reloadtime=0
	pages.strPageVar = "page"
	pages.InitClass
	Arr_Rs = pages.arrRecordInfo
	PageInfo = pages.strPageInfo
	page=pages.intPageNow
	Set pages = nothing
	Template=YxBBs.Template.ReadTemplate("������Ϣ")
	
	If IsArray(Arr_Rs) Then
		Lou = (Page-1)*10
		For i = 0 to UBound(Arr_Rs,2)
			UserName=Arr_Rs(10,i)
			BBSID=Arr_Rs(0,i)
			YxBBs_ubb.UbbString=Arr_Rs(7,i)
			Temp=Replace(Template,"{��Ϣ}","")
			Temp=Replace(Temp,"{����}","")
			Temp=Replace(Temp,"{QQ}",YxBBs.Fun.GetSqlStr(Arr_Rs(12,i)))
			Temp=Replace(Temp,"{����}",YxBBs.Fun.GetSqlStr(Arr_Rs(18,i)))
			Temp=Replace(Temp,"{��ҳ}",YxBBs.Fun.GetSqlStr(Arr_Rs(19,i)))
			Temp=Replace(Temp,"{�༭}","")
			Temp=Replace(Temp,"{����}","")
			Temp=Replace(Temp,"{¥��}","<b><font color=#FF0000>"&Lou+i+1&"</font><sup>F</sup></b>")
			dim ad
			SET ad=YxBBs.Execute("Select google From [YX_config]")
			If Arr_Rs(1,i)=ID Then
				Temp=Replace(Temp,"{ɾ��}",SetTopic())
				Temp=Replace(Temp,"{���}",ad(0))
				ShowCaption=TopicCaption
				ShowMood=TopicMood()
			Else
				ShowCaption="":ShowMood=""
				Temp=Replace(Temp,"{���}","")
				Temp=Replace(Temp,"{ɾ��}","<a title=ɾ���˻ظ� href=admin.Asp?BoardID={���ID}&ID={����ID}&BbsID={����ID}&TB={���ݱ�ID}&Action=ɾ���ظ�&>��</a>")
			End If
			Temp=Replace(Temp,"{���ݱ�ID}",YxBBs.TB)
			Temp=Replace(Temp,"{���ID}",YxBBs.BoardID)
			Temp=Replace(Temp,"{����ID}",ID)
			Temp=Replace(Temp,"{ҳ��}",page)
			Temp=Replace(Temp,"{����ID}",Arr_Rs(0,i))
			Temp=Replace(Temp,"{�û�����}",Arr_Rs(10,i))
			Temp=Replace(Temp,"{����ʱ��}",Arr_Rs(5,i))
			SET Rs=YxBBs.Execute("Select GradeName,GradePic From [YX_UserGrade] where GradeNum="&Arr_Rs(16,i)&"")
            Temp=Replace(Temp,"{�ȼ�ͼƬ}","<img src='"&Rs("Gradepic")&"' alt="&Rs("GradeName")&" Align=center>")
			Temp=Replace(Temp,"{�ȼ�����}",Rs("GradeName"))
			IF YxBBs.ClassID>2 then
				Temp=Replace(Temp,"{�û�IP}",YxBBs.Fun.GetIp(Arr_Rs(8,i),2))
			Else
				Temp=Replace(Temp,"{�û�IP}",YxBBs.Fun.GetSqlStr(Arr_Rs(8,i)))
			End If
			Temp=Replace(Temp,"{Y��}",""&YxBBs.BBSSetting(23)&"��<b><font color=red>"&YxBBs.Fun.GetSqlStr(Arr_Rs(21,i))&"</font></b>")
			Temp=Replace(Temp,"{��Ǯ��}",YxBBs.Fun.GetSqlStr(Arr_Rs(22,i)))
			
			If Arr_Rs(29,i)<>"" Then
				Temp=Replace(Temp,"{ͷ��}","ͷ�Σ�"&YxBBs.Fun.GetSqlStr(Arr_Rs(29,i))&"<br>")
			Else
				Temp=Replace(Temp,"{ͷ��}","ͷ�Σ�δ����<br>")
			End If
			If IsOnline(Arr_Rs(10,i)) Then
				If Arr_Rs(20,i) Then
					Temp=Replace(Temp,"{�Ա�}","<img src='Images/11.gif' alt='�Ա���<br>״̬������<br>��ʱ��" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				Else
					Temp=Replace(Temp,"{�Ա�}","<img src='Images/01.gif' alt='�Ա�Ů<br>״̬������<br>��ʱ��" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				End If
			Else
				If Arr_Rs(20,i) Then
					Temp=Replace(Temp,"{�Ա�}","<img src='Images/10.gif' alt='�Ա���<br>״̬������<br>��ʱ��" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				Else
					Temp=Replace(Temp,"{�Ա�}","<img src='Images/00.gif' alt='�Ա�Ů<br>״̬������<br>��ʱ��" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				End If
			End If
			Temp=Replace(Temp,"{����}",Arr_Rs(17,i))
			Temp=Replace(Temp,"{ע��ʱ��}",formatdatetime(Arr_Rs(24,i),2))
			IF Arr_Rs(11,i) then
				Temp=Replace(Temp,"{�û�ͷ��}","<img src='http://qqshow-user.tencent.com/"&Arr_Rs(12,i)&"/10/'>")
			else
				Temp=Replace(Temp,"{�û�ͷ��}","<img border='0' src='"&Arr_Rs(13,i)&"' width="&Arr_Rs(14,i)&" height="&Arr_Rs(15,i)&">")
			End If
			Sign=Arr_Rs(23,i)
			IF Arr_Rs(25,i) then
				Temp=Replace(Temp,"{��������}","<Hr Size=1><Font Color=red>���û������ѱ�����Ա���Σ���͹���Ա��ϵ</Font><Hr Size=1>")
				Temp=Replace(Temp,"{ǩ��}","��ǩ�������ѹ���Ա���Ρ�")
			Else
				IF YxBBs.BBSSetting(12)="1" then Sign="" 
				IF Arr_Rs(26,i) Then Sign="��ǩ�����ѱ�����Ա���Ρ�"
				IF isNull(sign) or sign="" then
					Temp=Replace(Temp,"{ǩ��}",YxBBs.BBSName&"��ӭ����")
				Else
					Temp=Replace(Temp,"{ǩ��}",YxBBs_ubb.Sign_Code(Sign))
				End IF
				Temp=Replace(Temp,"{��������}",ShowMood&"<SPAN style='line-height:150%;table-layout:fixed;word-wrap:break-word;word-break:break-all;width:100%'><b>"&ShowCaption&"</b><br><br>"&YxBBs_ubb.UBB(Arr_Rs(3,i),1)&"</SPAN>")
			End IF
			TempStr=TempStr&Temp
		Next
		Response.Write(TempStr)
	End If
End Function

Function IsOnline(TUserName)
	Dim Temp,Rs,Ary_Rs,i
	IsOnline=false
 		Cache.Name="UserOnline"
	If Cache.valid Then
	   Temp=Cache.Value
	Else
		Set Rs=YxBBs.Execute("Select Name From [YX_online] where ClassID<>6")
		If Rs.Eof Or Rs.Bof Then Exit Function
		Ary_Rs=Rs.GetRows(-1)
		Rs.Close
		For i=0 To Ubound(Ary_Rs,2)
			Temp=Temp & Ary_Rs(0,i)&","
		Next
		Temp=","&Temp
		Cache.add Temp,dateadd("n",5,YxBBs.NowBBSTime)'5����
	End If
	If InStr(Lcase(Temp),","&Lcase(TUserName)&",")>0 Then IsOnline=true
End Function

Function SetTopic()
	Dim Temp,Url
	If Not YxBBs.FoundUser Then Exit Function
	Url="admin.Asp?BoardID="&YxBBs.BoardID&"&ID="&ID&"&TB="&YxBBs.TB&"&Action="
        
	IF TopicTopType=5 then Temp="<div class=menuitems><a href="&Url&"ȡ�����ö�>ȡ�����ö�</a></div>" Else Temp="<div class=menuitems><a href="&Url&"���ö�>��Ϊ���ö�</a></div>"
	IF TopicTopType=4 Then Temp=Temp&"<div class=menuitems><a href="&Url&"ȡ�����ö�>ȡ�����ö�</a></div>" ELse Temp=Temp&"<div class=menuitems><a href="&Url&"���ö�>��Ϊ���ö�</a></div>"
	IF TopicTopType=3 Then Temp=Temp&"<div class=menuitems><a href="&Url&"ȡ���ö�>ȡ���ö�</a></div>" ELse Temp=Temp&"<div class=menuitems><a href="&Url&"�ö�>��Ϊ�ö�</a></div>"
	IF TopicIsGood Then Temp=Temp&"<div class=menuitems><a href="&Url&"ȡ������>ȡ������</a></div>" Else Temp=Temp&"<div class=menuitems><a href="&Url&"����>��Ϊ����</a></div>"
	IF TopicIsLock Then Temp=Temp&"<div class=menuitems><a href="&Url&"����>�������</a></div>" Else Temp=Temp&"<div class=menuitems><a href="&Url&"����>��������</a></div>"
	Temp=Temp&"<div class=menuitems><a href="&Url&"ɾ������>ɾ������</a></div><div class=menuitems><a href="&Url&"�ƶ�>�ƶ�����</a></div><div class=menuitems><a href="&Url&"����>��������</a></div>"
       Temp="<a onmouseover=""showmenu(event,'"&Temp&"')"">����</a>"
	SetTopic=Temp
End Function

Function FastReply()
	If Not YxBBs.FoundUser Then Exit Function
	IF Not TopicIsLock Or YxBBs.ClassID=1 then
		Dim Caption,Content
		Caption="���ٻظ���<a href='Say.Asp?action=reply&BoardID="&YxBBs.BoardID&"&id="&ID&"'>�߼��ظ�</a>��"
		Content="<form style='margin: 8;' method='POST' name='yimxu' onsubmit='if(content.value==""""){alert(""����û����д��������!"");content.focus();return (false);};"
		If Cint(YxBBs.BoardSetting(0))=1 then Content=Content&"if(CheckCode.value==""""){alert(""����д��֤�룡"");CheckCode.focus();return (false);};"
		Content=Content&"BtnPost.disabled = true;' action='Save.Asp?action=Reply&BoardID="&YxBBs.BoardID&"&TB="&YxBBs.TB&"&ID="&ID&"&page="&Page&"'><table align='center' border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='100%'><tr><td width='15%' valign='top'><p style='line-height: 150%; margin: 6'>�ظ�˵����<br>��̳���Ҽ�,ά�������,�����ع��ҷ���,���ҷ�ˮ��!</p>"&_
		"<td width='76%' valign='top'><input type=hidden name='caption' value='Re:"&TopicCaption&"'><textarea onkeydown=presskey(); rows=6 name=content Style='width:98%'></textarea><p style='margin-top: 10; margin-bottom: 0'>"
		IF Cint(YxBBs.BoardSetting(0))=1 then Content=Content&""&YxBBs.GCode&""
		Content=Content&"<input type=submit value='�ظ�����' ID='BtnPost'> <input type='reset' value='�� д'> ��ݼ� Ctrl+Enter </p></td><td width='5%' valign='top'></td></tr></table></form>"
	Call YxBBs.ShowTable(Caption,Content)
	END IF
End Function
%>