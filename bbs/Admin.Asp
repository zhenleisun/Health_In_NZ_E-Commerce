<!--#include file="Inc/SysConfig.Asp"-->
<%
Dim Caption,Content,Rs,ID,Url,Action,SetUserName,GoToUrl,BoardID
YxBBs.Fun.CheckMake()
ID=YxBBs.Checknum(Request.QueryString("ID"))
Action=lcase(Request.QueryString("Action"))
YxBBs.CheckBoard()
URl="Show.Asp?ID="&ID
GoToUrl=True
Call YxBBs.Head("��̳����")
Select Case Action
Case"����"
Call cheak()
	If Cint(YxBBs.ClassSetting(19))=0 Then YxBBs.Error("��û�в�����������Ȩ�ޣ�")
	Set Rs=YxBBs.Execute("Select IsGood,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		IF Rs(0) Then
			YxBBs.Error("�����������Ѿ��Ǿ��������ˣ�")
		ELse
			YxBBs.Execute("update [YX_Topic] set IsGood=True where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.Execute("update [YX_User] set Coin=Coin+"&YxBBs.BBSSetting(39)&",Mark=Mark+"&YxBBs.BBSSetting(40)&",GoodNum=GoodNum+1 where name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"����Ϊ����������ID��"&ID,Rs(1),0
		End If
	End If
	Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�趨Ϊ�������ӳɹ�!3��󷵻�����!"
	Rs.Close
	Call Show()
Case"ȡ������"
Call cheak()
	If Cint(YxBBs.ClassSetting(19))=0 Then YxBBs.Error("��û�в�����������Ȩ�ޣ�")
	Set Rs=YxBBs.Execute("Select IsGood,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.Eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		IF Not Rs(0) Then
			YxBBs.Error("���������Ӳ��Ǿ������ӣ�")
		ELse
			YxBBs.Execute("update [YX_Topic] set IsGood=False where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_User] set Coin=Coin-"&YxBBs.BBSSetting(41)&",Mark=Mark-"&YxBBs.BBSSetting(42)&",GoodNum=GoodNum-1 where name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"��ȡ������������ID��"&ID,Rs(1),0
		End If
	End If
	Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">ȡ�����Ӿ����ɹ�!3��󷵻�����!"
	Rs.Close
	Call Show()
Case"�ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(16))=0 Then YxBBs.Error("��û���ö������Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		IF Rs(0)=5 Then
			YxBBs.Error("�������Ѿ������ö�����")
		ElseIf Rs(0)=4 Then
			YxBBs.Error("�������Ѿ������ö�����")
		ElseIF Rs(0)=3 Then
			YxBBs.Error("�����������Ѿ����ö��ˣ�")	
		Else
			YxBBs.Execute("update [YX_Topic] Set TopType=3 where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_User] set Coin=Coin+"&YxBBs.BBSSetting(43)&",Mark=Mark+"&YxBBs.BBSSetting(44)&" Where name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"����Ϊ�ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ���������Ϊ�ö�����!3��󷵻�����!"
	End If
	Rs.Close
	Call Show()
Case"ȡ���ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(16))=0 Then YxBBs.Error("��û�в����ö������Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		IF Rs(0)=5 Then
			YxBBs.Error("�������Ѿ������ö�����")
		ElseIf Rs(0)=4 Then
			YxBBs.Error("�������Ѿ������ö�����")
		ElseIF Rs(0)<>3 Then
			YxBBs.Error("�������Ѿ������ö�����")
		Else
			YxBBs.Execute("update [YX_Topic] Set TopType=0 where TopicId="&ID&" ")
			YxBBs.execute("update [YX_User] set Coin=Coin-"&YxBBs.BBSSetting(45)&",Mark=Mark-"&YxBBs.BBSSetting(46)&" Where Name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"��ȡ���ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ���ȡ�����ӵ��ö�!3��󷵻�����!"
	End If
	Set Rs=Nothing
	Call Show()
Case"���ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(17))=0 Then YxBBs.Error("��û�й������ö����ӵ�Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And (BoardID="&YxBBs.BoardID&" or TopType=5)")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Rs(0)=5 Then
			YxBBs.Error("�������Ѿ������ö����⣡")
		Else
			YxBBs.execute("update [YX_Topic] Set TopType=5 where TopicID="&id&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_user] Set Coin=Coin+"&YxBBs.BBSSetting(51)&",Mark=Mark+"&YxBBs.BBSSetting(52)&" where Name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"����Ϊ���ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ�����������Ϊ���ö�!3��󷵻�����!"
	End If
	Rs.Close
	Call Show()
Case"ȡ�����ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(17))=0 Then YxBBs.Error("��û�й������ö����ӵ�Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And (BoardID="&YxBBs.BoardID&" or TopType=5)")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Rs(0)<>5 Then
			YxBBs.Error("�����Ӳ������ö����⣡")
		Else
		YxBBs.execute("update [YX_Topic] set TopType=0 where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
		YxBBs.execute("update [YX_user] set Coin=Coin-"&YxBBs.BBSSetting(53)&",Mark=Mark-"&YxBBs.BBSSetting(54)&" where name='"&SetUserName&"'")
		YxBBs.InLog "�����⡶"&Rs(2)&"��ȡ�����ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ���ȡ�����ӵ����ö�!3��󷵻�����!"
	End If
	Rs.Close
	Call Show()
Case"���ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(17))=0 Then YxBBs.Error("��û�й������ö����ӵ�Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And (BoardID="&YxBBs.BoardID&" or TopType=4)")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Rs(0)=5 Then
			YxBBs.Error("�������Ѿ������ö����⣡")
		ElseIf Rs(0)=4 Then
			YxBBs.Error("�������Ѿ������ö����⣡")
		Else
			YxBBs.execute("update [YX_Topic] Set TopType=4 where TopicID="&id&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_user] Set Coin=Coin+"&YxBBs.BBSSetting(47)&",Mark=Mark+"&YxBBs.BBSSetting(48)&" where Name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"����Ϊ���ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ�����������Ϊ���ö�!3��󷵻�����!"
	End If
	Rs.Close	
	Call Show()
Case"ȡ�����ö�"
Call cheak()
	If Cint(YxBBs.ClassSetting(17))=0 Then YxBBs.Error("�Բ�����û�й���Ȩ�ޣ�")
	IF YxBBs.ClassID>2 Then YxBBs.Error("�Բ�����û�й���Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select TopType,Name,Caption From[YX_Topic] where TopicID="&ID&" And (BoardID="&YxBBs.BoardID&" or TopType=4)")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Rs(0)<>4 Then
			YxBBs.Error("�����Ӳ������ö����⣡")
		Else
			YxBBs.execute("update [YX_Topic] set TopType=0 where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_user] set Coin=Coin-"&YxBBs.BBSSetting(49)&",Mark=Mark-"&YxBBs.BBSSetting(50)&" where name='"&SetUserName&"'")
			YxBBs.InLog "�����⡶"&Rs(2)&"��ȡ�����ö�������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�ɹ���ȡ�����ӵ����ö�!3��󷵻�����!"
	End If
	Rs.Close
	Call Show()
Case"����"
Call cheak()
	If Cint(YxBBs.ClassSetting(20))=0 Then YxBBs.Error("��û�����������Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select IsLock,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Rs(0) Then
			YxBBs.Error("�����������Ѿ���������")
		Else
			YxBBs.execute("update [YX_Topic] set IsLock=True where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.InLog "�����⡶"&Rs(2)&"������������ID��"&ID,Rs(1),0
		End If
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�����Ѿ������ɹ�!3��󷵻�����!"
	End If
	Rs.Close			
	Call Show()
Case"����"
Call cheak()
	If Cint(YxBBs.ClassSetting(20))=0 Then YxBBs.Error("��û�н��������Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select Islock,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	If Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Not Rs(0) Then
			YxBBs.Error("������û�б�������")
		Else
			YxBBs.execute("update [YX_Topic] set IsLock=False where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.InLog "�����⡶"&Rs(2)&"���������������ID��"&ID,Rs(1),0
		End IF
		Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">������������ɹ�!3��󷵻�����!"
	End if
	Rs.Close	
	Call Show()

Case"ɾ������"
Call cheak()
	If Cint(YxBBs.ClassSetting(14))=0 Then YxBBs.Error("��û��ɾ�������Ȩ�ޣ�")
Set Rs=YxBBs.execute("Select IsDel,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	IF Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
End If
	Caption="ɾ������"
	Content="<form method=POST name='move' action='?BoardID="&YxBBs.BoardID&"&ID="&ID&"&TB="&YxBBs.TB&"&Action=ɾ������1'> <input type=submit value=' ȷ��ɾ�� '> <br><br>�Ƿ�������֪ͨ���ӵ����ߣ�<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.value=""֪ͨ���������ӡ�"&Rs(2)&"��������Ա("&YxBBs.MyName&")ɾ����ԭ���ǣ�""}else{sms.disabled=true;sms.value="""";}' type='checkbox' value='yes'> <input name='sms' size='50' type='text' value='' disabled='true'></p></form>"
	YxBBs.ShowTable Caption,"<tr><td height=""100"">"&Content&"</td></tr>"
Case"ɾ������1"
Call cheak()
        IsSms=YxBBs.Fun.GetStr("issms")
	Sms=YxBBs.Fun.GetStr("sms")
	If Cint(YxBBs.ClassSetting(14))=0 Then YxBBs.Error("��û��ɾ�������Ȩ�ޣ�")
	Set Rs=YxBBs.execute("Select IsDel,Name,Caption From[YX_Topic] where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	IF Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Not Rs(0) Then
			GoToUrl=False
			'�ó�Ҫɾ��������
			Temp=YxBBs.Execute("Select Count(BbsID) From[YX_bbs"&YxBBs.TB&"] where (TopicID="&ID&" Or  ReplyTopicID="&ID&") And BoardID="&YxBBs.BoardID)(0)
			YxBBs.Execute("update [YX_Topic] set IsDel=True,ReplyNum=0 where TopicID="&ID&" And BoardID="&YxBBs.BoardID)
			YxBBs.Execute("update [YX_Bbs"&YxBBs.TB&"] set IsDel=True where (TopicID="&ID&" Or  ReplyTopicID="&ID&") And BoardID="&YxBBs.BoardID) 
			YxBBs.Execute("update [YX_Config] set AllEssayNum=AllEssayNum-"&Temp&",TopicNum=TopicNum-1")
                        YxBBs.Execute("update [YX_User] set Coin=Coin-"&YxBBs.BBSSetting(59)&",Mark=Mark-"&YxBBs.BBSSetting(60)&" Where Name='"&SetUserName&"'")			
                        YxBBs.Execute("update [YX_Board] set EssayNum=EssayNum-"&Temp&",TopicNum=TopicNum-1 where BoardID="&YxBBs.BoardID&"")
                   If IsSms="yes" Then
                YxBBs.Execute("insert into [YX_Sms](name,MyName,Title,Content) values('ϵͳ��Ϣ','"&SetUserName&"','�������ⱻɾ��֪ͨ','"&Sms&"')")
		YxBBs.Execute("update [YX_User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+"&Len(Sms)&" Where Name='"&SetUserName&"'")
                End if
			If YxBBs.BoardDepth>1 Then
				YxBBs.Execute("Update [YX_Board] set TopicNum=TopicNum-1,EssayNum=EssayNum-"&Temp&" where BoardID In ("&YxBBs.BoardParentStr&") And Depth>0")
			End If
			BoardLastRaply()
			YxBBs.InLog "�����⡶"&Rs(2)&"��ɾ��������ID��"&ID,Rs(1),0
		Else
			YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
		End IF
		Caption="<meta http-equiv=refresh content=3;url=List.Asp?BoardID="&YxBBs.BoardID&">�ɹ�ɾ���������ӣ�"
	End if
	Rs.Close
	Call Show()
Case"ɾ���ظ�"
Call cheak()
	If Cint(YxBBs.ClassSetting(14))=0 Then YxBBs.Error("��û��ɾ���ظ���Ȩ�ޣ�")
	GoToUrl=False
	Dim BbsID,Temp
	BbsID=YxBBs.CheckNum(Request.QueryString("BbsID"))
	Set Rs=YxBBs.execute("Select IsDel,Name,Caption From [YX_bbs"&YxBBs.TB&"] where ReplyTopicID="&ID&" And BbsID="&BbsID&" And BoardID="&YxBBs.BoardID&"")
	IF Rs.eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
	Else
		IF Not Rs(0) Then
			YxBBs.execute("update [YX_bbs"&YxBBs.TB&"] set IsDel=True where ReplyTopicID="&ID&" And BbsID="&BbsID&" And BoardID="&YxBBs.BoardID&"")
			YxBBs.execute("update [YX_Config] Set AllEssayNum=AllEssayNum-1")
			YxBBs.execute("update [YX_Board] Set EssayNum=EssayNum-1 where BoardID="&YxBBs.BoardID&"")
                        YxBBs.execute("update [YX_Topic] set ReplyNum=ReplyNum-1 where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
			If YxBBs.BoardDepth>1 Then
				YxBBs.Execute("Update [YX_Board] set TopicNum=TopicNum-1,,EssayNum=EssayNum-1 where BoardID In ("&YxBBs.BoardParentStr&") And Depth>0")
			End If
			BoardLastRaply()
			YxBBs.InLog "ɾ���ظ���"&Rs(2)&"��������ID��"&ID,Rs(1),0
		Else
			YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ��")
		End IF
		Caption="�ɹ�ɾ���ظ����ӣ�"
	End if
	Rs.Close
	Call Show()
Case"�ƶ�"
Call cheak()
	If Cint(YxBBs.ClassSetting(15))=0 Then YxBBs.Error("��û���ƶ������Ȩ�ޣ�")
	GoToUrl=False
	Caption="�ƶ�����"
	Content="<form method=POST name='move' action='?action=move&TB="&YxBBs.TB&"&ID="&ID&"&BoardID="&YxBBs.BoardID&"'><p style='margin: 15'>��ѡ������Ҫ�ƶ�������̳��"&GetBoardList()&" <input type=submit value=' �� �� '> <br><br>�Ƿ�������֪ͨ���ӵ����ߣ�<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.value=""֪ͨ���������ӱ�����Ա("&YxBBs.MyName&")�ƶ��ˣ�""}else{sms.disabled=true;sms.value="""";}' type='checkbox' value='yes'> <input name='sms' size='50' type='text' value='' disabled='true'></p></form>"
	YxBBs.ShowTable Caption,"<tr><td height=""100"">"&Content&"</td></tr>"
Case"move"
Call cheak()
	If Cint(YxBBs.ClassSetting(15))=0 Then YxBBs.Error("��û���ƶ������Ȩ�ޣ�")
	Dim IsSms,Sms,TmpBoardID
	GoToUrl=False
	IsSms=YxBBs.Fun.GetStr("issms")
	Sms=YxBBs.Fun.GetStr("sms")
	TmpBoardID=YxBBs.Checknum(Request.form("BoardID"))
	Set Rs=YxBBs.Execute("Select BoardID,Name,Caption From[YX_Topic] where TopicID="&ID&" and BoardID="&YxBBs.BoardID&"")
	YxBBs.execute("update [YX_Topic] Set BoardID="&TmpBoardID&" where TopicID="&ID&"")
	YxBBs.execute("update [YX_bbs"&YxBBs.TB&"] Set BoardID="&TmpBoardID&" where TopicId="&ID&" or ReplyTopicid="&ID&"")
	YxBBs.InLog "�����⡶"&Rs(2)&"����IDΪ"&Rs(0)&"�İ����ƶ�������IDΪ"&TmpBoardID&"�İ��棬����ID��"&ID,Rs(1),0
	If IsSms="yes" Then
		Sms=Sms&vbcrlf&"���⣺"&Rs(2)&",���ƶ�������IDΪ"&TmpBoardID&"�İ��."
		YxBBs.Execute("insert into [YX_Sms](name,MyName,Title,Content) values('ϵͳ��Ϣ','"&SetUserName&"','�������ⱻ�ƶ�֪ͨ','"&Sms&"')")
		YxBBs.Execute("update [YX_User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+"&Len(Sms)&" Where Name='"&SetUserName&"'")
	End If
	Rs.Close
	Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">�����ƶ��ɹ�!3��󷵻�����!"
	Call Show()
Case"����"
Call cheak()
       If Cint(YxBBs.ClassSetting(21))=0 Then YxBBs.Error("��û�����������Ȩ�ޣ�")
	YxBBs.Execute("update [YX_Topic] set LastTime='"&YxBBs.NowBbsTime&"' where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	YxBBs.Execute("update [YX_bbs"&YxBBs.TB&"] set LastTime='"&YxBBs.NowBbsTime&"' where TopicID="&ID&" And BoardID="&YxBBs.BoardID&"")
	Set Rs=YxBBs.Execute("Select BoardID,Name,Caption From[YX_Topic] where TopicID="&ID&" and BoardID="&YxBBs.BoardID&"")
	YxBBs.InLog "�ɹ��������⡶"&Rs(2)&"��������ID��"&ID,Rs(1),0
	Rs.Close
	Caption="<meta http-equiv=refresh content=3;url=Show.Asp?ID="&ID&">���������ɹ�!3��󷵻�����!"
	Call Show()

Case "editsetting"
Dim BoardSetting,sql,Brule
IF YxBBs.ClassID>3 Then YxBBs.Error("����Ȩ�޲��㣡")
BoardID=Request.form("BoardID")
Brule=YxBBs.Fun.GetStr("Brule")
IF BoardID="" or not isNumeric(BoardID) then
	YxBBs.Error("���������룬��ȷ����ȷ�Ĳ�����")
End if
BoardSetting=Request.form("TopicCode") & "," & Request.form("UpLoadType") & "," & Request.form("TopicMaxLong") & "," & Request.form("ArticleMaxLong") & "," & Request.form("BoardMaxVote") & "," & Request.form("UserMinPost") & "," & Request.form("UserMinRemark") & "," & Request.form("UserMinMoneys") & "," & Request.form("UserMinLoginTimes") & "," & Request.form("UserMinDigestNum") & "," & Request.form("UserMinRegTime") & "," & Request.form("ALeast")
Set Rs=Server.CreateObject("Adodb.Recordset")
YxBBs.execute("update [YX_Board] Set BoardSetting='"&BoardSetting&"',Brule='"&Brule&"' where BoardID="&BoardID&"")
Set rs=nothing
Cache.name="BoardInfo":Cache.clean()
Caption="�����ɹ�"
Content="<div style='margin:15;line-height: 150%'><li>���Ȩ�����óɹ��� <li><a href=?BoardID="&BoardID&">���ذ������</a></div>"
Call YxBBs.ShowTable(Caption,Content)


Case Else
IF YxBBs.ClassID>3 Then YxBBs.Error("����Ȩ�޲��㣡")
IF  YxBBs.ClassID=3 And YxBBs.IsBoardAdmin=False Then YxBBs.Error("����Ȩ�޲��㣡")
dim ClassID,BoardName
BoardID=Request("BoardID")
if Not IsNumeric(ClassID) then
	YxBBs.Error("������������¼�飡")
end if

set rs=Conn.execute("Select BoardID,BoardName,BoardSetting,brule from YX_Board where BoardID="& BoardID)
if rs.eof or rs.bof then
	YxBBs.Error("�Ҳ����˰�����ɾ����")
else
	BoardName=rs("BoardName")
	BoardSetting=split(rs("BoardSetting"),",")
Temp=YxBBs.Template.ReadTemplate("�����б�����")
	Temp=Replace(Temp,"{����}",YxBBs.Stats )
	Response.Write Temp
%>
<form action='?action=editsetting&BoardID=<%=BoardID%>' method=post name='creator'><table cellpadding=6 cellspacing=1 border=0 width=100% align='center'>
 <tr><td width=35% height=20>�������</td><td height=20><INPUT TYPE=hidden name=BoardID value=<%=BoardID%> readonly><%=BoardName%></td>
</tr><tr><td colspan=2 height=14><font color=#000077><b>������</b></font></td>
</tr><tr><td height=20>����ʹ����֤��</td>
<td height=20>��<input name=TopicCode type=radio value=1
                <%if BoardSetting(0)=1 then%>
                checked>
                <%else%>
                >
		<%end if%>
&nbsp;��<input name=TopicCode type=radio value=0
		<%if BoardSetting(0)=0 then%>
checked></td>
                <%else%>
></td>
                <%end if%>
</tr><tr><td height=20>�ϴ�����</td><td height=0>
<input name=UpLoadType size=20 value=<%=BoardSetting(1)%>></td></tr><tr><td height=20>�������Ƴ���</td>
<td height=20><input name=TopicMaxLong size=20 value=<%=BoardSetting(2)%>></td></tr><tr>
<td height=26>������������ֽ���</td><td height=26><input name=ArticleMaxlong size=20 value=<%=BoardSetting(3)%>></td>
</tr><tr><td height=20>���ͶƱ��Ŀ</td><td height=20>
<input name=BoardMaxVote size=20 value=<%=BoardSetting(4)%> ></td></tr><tr>
</tr><tr><td>����������<br>* ���ð�������֧��HTML.<br>* ���Ĭ��֧����ʽ<font color=red>&lt;div class=rule&gt;�������&lt;/div&gt;</font><br>* ����ֻ֧��255���ַ�.  </td><td>
<textarea rows=6 cols=70 name=Brule><%=Rs("Brule")%></textarea></td></tr><tr>
<td colspan=2 height=14><font color=#000077><b>�û����</b></font></td>		
</tr><tr><td height=20>�û�����������</td><td height=20>
<input name=UserMinPost size=20 value=<%=BoardSetting(5)%> ></td></tr><tr>
<td height=20>�û����ٻ���</td><td height=20>
<input name=UserMinRemark size=20 value=<%=BoardSetting(6)%> ></td>
</tr><tr><td height=20>�û����ٽ�Ǯ</td><td height=20>
<input name=UserMinMoneys size=20 value=<%=BoardSetting(7)%> ></td></tr><tr>
<td height=20>��½�������ټ���</td><td height=20>
<input name=UserMinLoginTimes size=20 value=<%=BoardSetting(8)%> ></td>
</tr><tr><td height=20>�û����پ�������</td><td height=20>
<input name=UserMinDigestNum size=20 value=<%=BoardSetting(9)%> ></td>
</tr><tr><td height=20>����ע��ʱ�䣨��λΪ���ӣ�</td><td  height=20>
<input name=UserMinRegTime size=20 value=<%=BoardSetting(10)%> ></td></tr><tr>
<td height=26>�������������ֽ���</td><td height=26><input name=ALeast size=20 value=<%=BoardSetting(11)%> ></td></tr><tr>
<td valign=middle colspan=2 align=center height=20><input type=submit value='�� ��' name='B1'></td></tr>
</table></form>
<%	
        Temp=YxBBs.Template.ReadTemplate("�����б��ײ�")
	Temp=Replace(Temp,"{��ҳ}","")
	Response.Write Temp
end if
End Select
YxBBs.Footer()
Set YxBBs=Nothing

Sub Show()
	Dim Temp
	IF GoToUrl Then Temp="<li><a href="&Url&">�ص�����</a></li>"
	Content=Content&Temp&"<li><a href=List.Asp?BoardID="&YxBBs.BoardID&">"&YxBBs.BoardName&"</a></li><li><a href=Default.Asp>������ҳ</a>"
	YxBBs.Success Caption,Content
End Sub

Function GetBoardList()
	Dim Temp,I,II,po
	Temp="<select Style='font-size: 9pt' name='BoardID' >"
	If IsArray(YxBBs.Board_Rs) Then
		For i=0 To Ubound(YxBBs.Board_Rs,2)
			Po=""
			If YxBBs.Board_Rs(0,i)=0 Then
				Temp=Temp&"<option value='"&YxBBs.Board_Rs(1,i)&"'>��"&YxBBs.Board_Rs(3,i)&"��</option>"
			Else
				For II=2 to YxBBs.Board_Rs(0,i)
					po=Po&"�O"
				Next
				Temp=Temp&"<option value='"&YxBBs.Board_Rs(1,i)&"'>"&po&"�� "&YxBBs.Board_Rs(3,i)&"</option>"
			End IF
		Next
	End If
	GetBoardList=Temp&"</select>"
End Function

Sub BoardLastRaply()
	Dim Rs,Temp,TempID,TempContent
	Set Rs=YxBBs.execute("select top 1 ReplyTopicID,Name,Caption,Content,TopicID,LastTime,BoardID from [YX_bbs"&YxBBs.TB&"] where BoardID="&YxBBs.BoardID&" And IsDel=False order by BbsID desc")
	if Rs.eof then
		Temp=""
	Else
		if Rs("ReplyTopicID")=0 then
			TempContent=YxBBs.Fun.StrLeft(Rs("Caption"),20)
			TempID=Rs("Topicid")
		Else
			TempContent=YxBBs.Fun.StrLeft(YxBBs.Fun.FixReply(Rs("Caption")),20)
			TempID=Rs("ReplyTopicID")
			'��������
			Dim ReRs,TopicLastReply
			Set ReRs=YxBBs.execute("select top 1 Name,Content from [YX_bbs"&YxBBs.TB&"] where BoardID="&YxBBs.BoardID&" And ReplyTopicID="&TempID&" And IsDel=False order by BbsID desc")
			If Not ReRs.Eof Then
				TopicLastReply=""&ReRs(0)&"|"&YxBBs.Fun.StrLeft(ReRs(1),40)
				YxBBs.execute("Update [YX_Topic] set LastReply='"&TopicLastReply&"' where TopicId="&TempID&"")
			End If
			ReRs.CLose:Set ReRs=Nothing
		End if
	Temp=""&Rs("name")&"|"&TempContent&"|"&Rs("LastTime")&"|99|"&TempID&"|"&Rs("BoardID")&"|"&YxBBs.TB&""
	end if
	Rs.Close
	IF Temp<>"" Then 
		YxBBs.execute("update [YX_Board] set LastReply='"&Temp&"' where BoardID="&YxBBs.BoardID&" and ParentID<>0")
		If YxBBs.BoardDepth>1 Then
			YxBBs.Execute("Update [YX_Board] set LastReply='"&Temp&"' where BoardID In ("&YxBBs.BoardParentStr&") And ParentID<>0")
		End If
	End IF
	Cache.name="Config"
	Cache.clean()
	Cache.Name="BoardInfo"
	Cache.clean()
End Sub

sub cheak
If YxBBs.Execute("Select ID From[YX_user]where Name='"&YxBBs.MyName&"' And PassWord='"&YxBBs.MyPwd&"'").Eof Then
	YxBBs.MakeCookiesEmpty()
	YxBBs.Error("���û������ڻ����Ѿ���ɾ����")
End IF

Set Rs=YxBBs.Execute("Select Name,BoardID From[YX_topic] Where TopicID="&ID&" And (BoardID="&YxBBs.BoardID&" or TopType=5)")
	IF Not Rs.eof Then
		SetUserName=Rs("Name")
		YxBBs.BoardID=Rs("BoardID")
		Rs.Close
	Else
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	End IF
end sub
%>