'============================================================================================================================================
'�\��G�C�g�@�۰�mail�U���@�H���L���Ӫ��D���ơA�D�n�O�D�ɪ����ʸ�T(�ͩR�i��)
'============================================================================================================================================
on error resume next		'��L�@��,���i�X��,�H�Khung��
MailTag="#{mail}:"
LeafTag="#{�w��¾}�G"
' �C�椣�[ & vbcrlf �|�X��
'--------------------------------��Ʈw�s��--------------------------------------------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=10.6.1.11;Trusted_Connection=True;Database=" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn & "IDMS" : conn.CommandTimeout = 1800
set rs=createobject("ADODB.Recordset")
set rs1=createobject("ADODB.Recordset")

set rs2=createobject("ADODB.Recordset")

set fs=createobject("scripting.filesystemobject")
set ff=fs.OpenTextFile("Mail_Life_" & DT(now,"yyyy") & ".log",8,true)
ff.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(Begin)------------------------------------------------------------------" & vbcrlf

set ff2=fs.OpenTextFile("Mail_Content_Dev.log",2,true)
ff2.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(Begin)------------------------------------------------------------------" & vbcrlf


call CheckDB()	'��s��Ƭd�ָ��

ChkMemo="---------- �d�ֶ��ػ����A�榡�G�ϥΤ����B�d�ֶ��ءB�d�ֻ��� ---------------------------------" & vbcrlf 
rs.open "select * from [Config] where [Kind]='�d�ֶ���' and [Item]<>'' and [Config] not in ('�n��D��','�n����v') order by [Mark]",conn
while not rs.eof
  ChkMemo=ChkMemo & rs(2) & "�@" & rs(1) & "�@" & rs(4) & vbcrlf
  rs.movenext
wend
ChkMemo=ChkMemo & "-------------------------------------------------------------------------------------" & vbcrlf
rs.close

rs1.open "select [����],[�Ƶ�] from [View_���u] where [���]='��T����' and [�ʽ�]<>'����' and [�Ƶ�] like '%" & MailTag & "%' and [�Ƶ�] not like '%" & LeafTag & "%'",conn	'�H��A�אּ�ۤv���ʪ������q��
while not rs1.eof
  '----------mail�}�Y��r------------------------------------
  LifeAll=rs1(0) & " ���x�z�n�G" & " " & vbcrlf & " " & vbcrlf
  LifeAll=LifeAll & "1. �귽�޲z�t��(IDMS)��@�g(" & DT(now-7,"yyyy/mm/dd") & " - " &  DT(now,"yyyy/mm/dd") & ")�d�֨ƶ��۰ʳq���p�U�ҥܡG" & vbcrlf 
  LifeAll=LifeAll & "2. �����d�߸��|�Ghttp://10.6.1.11:8888  -> [�귽�޲z]  -> [�ͩR�i��]" & vbcrlf & " " & vbcrlf & " " & vbcrlf

  flag=0	'�@�g�q���@���A�����ʰO���~�ݳq��
  ChkNum=0	'�|�ݧ󥿵���
  UpdNum=0	'���ʰO������

  PropNum=0	'�]���t�ε��ơA�Y<1000�h��mail[���`�L���]���s��]
  rs.open "select count(*) from [�]���D��]",conn
  if not rs.eof then PropNum=rs(0)
  rs.close

  '----------�d�ֳq���H��(�ư��Ȭ� �L�겣�s�� �� �����H��)-------------------------------------
  only_asset_related_flag=0
  
  SQL="SELECT COUNT(*) FROM GetCheckMaintainerAndRelatedToNotify() WHERE [�q���H��]='" + rs1(0) + "' OR [�q���H��] in (select Kind from Config where Item='" + rs1(0) + "')"
  rs.open SQL,conn  
  'ff.write "#### " & rs1(0) & " = " & rs(0) & " ####" & vbcrlf
  if CInt(rs(0))=0 then	
	'ff.write "SKIP!" & vbcrlf
	only_asset_related_flag=1
  end if
  rs.close
  
  '----------�d�ֳq���H��(�ư� �L�겣�s��)-------------------------------------
  flag2=0
  
  '----------�d�ָ�T-----------------------------------------
  SQL="select * from [��Ƭd��] where (" + GetSQL("[���@�H��]",rs1(0)) + " or " + GetSQL("[�����H��]",rs1(0)) + " or '" + rs1(0) + "' in (select [Item] from [Config] where [Kind]='SSM�p��') and [�d�ֵ��G] like '%�D�k%')"
  if PropNum<1000 then SQL=SQL & " and [�d�ֵ��G]<>'���`�L���]���s��'"
  SQL=SQL & " order by [���W��],[�D��s��]"
  rs.open SQL,conn
  if not rs.eof then
    LifeAll=LifeAll & "---------- ��Ƭd�֡G�ݧ󥿪����-------------------------------------------" & vbcrlf
    LifeAll=LifeAll & " �榡�G[�ϥΤ���]�@[�D��s��]�@[��Ƥ��e]�@[���@�H��]�@[�����H��]�@[�d�ֵ��G] " & vbcrlf
    LifeAll=LifeAll & "------------------------------------------------------------------------" & " " & vbcrlf
    while not rs.eof
      'if ChkNum<10 then LifeAll=LifeAll & rs(0) & "�@" & rs(1) & "�@" & rs(2) & "�@" & rs(3) & "�@" & rs(4) & " " & rs(5) & " "  & vbcrlf
      'ChkNum=ChkNum+1
	  
	  if rs(5)="�L�겣�s��" then		
		'rs2.open "SELECT * FROM Config WHERE Kind='" + rs(3) + "' AND Item='" + rs1(0) + "'",conn		
	    'if ((rs(3)=rs1(0)) or (not rs2.eof)) then 
	    '  if ChkNum<10 then LifeAll=LifeAll & rs(0) & "�@" & rs(1) & "�@" & rs(2) & "�@" & rs(3) & "�@" & rs(4) & " " & rs(5) & " "  & vbcrlf			
		'  ChkNum=ChkNum+1
		'end if 
		'rs2.close
	  else
		rs5_no_asset_no = rs(5)
		rs5_no_asset_no = Replace(rs5_no_asset_no, "�L�겣�s�� ", "")
		rs5_no_asset_no = Replace(rs5_no_asset_no, " �L�겣�s��", "")						
		
		'if ChkNum<10 then LifeAll=LifeAll & rs(0) & "�@" & rs(1) & "�@" & rs(2) & "�@" & rs(3) & "�@" & rs(4) & " " & rs(5) & " "  & vbcrlf					
		if ChkNum<10 then LifeAll=LifeAll & rs(0) & "�@" & rs(1) & "�@" & rs(2) & "�@" & rs(3) & "�@" & rs(4) & " " & rs5_no_asset_no & " "  & vbcrlf			
		ChkNum=ChkNum+1
		flag2=1
	  end if
		
      rs.movenext
    wend
    LifeAll=LifeAll & " " & vbcrlf
    flag=1
  end if
  rs.close
  if ChkNum>=10 then LifeAll=LifeAll & "********** �d�ָ�T" & ChkNum & "���A�W�L10�����������C�X�A�Цۦ�ܸ�Ƭd�֤����d�ߧ����� **********" & vbcrlf & vbcrlf
  if ChkNum>0 then LifeAll=LifeAll & ChkMemo & vbcrlf

  LifeAll=LifeAll & vbcrlf & vbcrlf
  '----------���ʰO��------------------------------------------
  LifeAll=LifeAll & "---------- �ͩR�i��(���ʰO��)�G�ȨѰѦҩνƮ� --------------------------------------------" & vbcrlf
  SQL=GetSQL("[���@�H��]",rs1(0)) + " or " + GetSQL("[��t�d�H]",rs1(0)) + " or " + GetSQL("[��O�ޤH]",rs1(0)) + " or [�D��s��] in (select [�]�ƽs��] from [View_�]�ƺ޲z] where [�]�ƽs��]<>null and [�O�ޤH��]='" & rs1(0) & "')"
  call GetLife("����]��",SQL)
  
  SQL="[�D��s��] in (select [�@�~�s��] from [View_�q�γ]��] where [�@�~�s��]<> null and (" + GetSQL("[�]�ƺ��@�H��]",rs1(0)) + " or [�O�ޤH��]='" + rs1(0) + "'))"
  call GetLife("�@�~�D��",GetSQL("[���@�H��]",rs1(0)) + "  or " + SQL)

  call GetLife("*",GetSQL("[���@�H��]",rs1(0)))   
  if UpdNum>=10 then LifeAll=LifeAll & "********** ���ʰO��" & UpdNum & "���A�W�L10�����������C�X�A�Цۦ�ܥͩR�i�������d�ߧ����� **********" & vbcrlf
  '----------mail�����H��--------------------------------------
  'if flag=1 and (only_asset_related_flag=0 or UpdNum<>0) then
  if flag=1 and ((only_asset_related_flag=0 and flag2=1) or UpdNum<>0) then
      who=rs1(0) : email=rs1(1)
      pos=instr(1,email,MailTag)
      if pos>0 then
        email=mid(email,pos+8,instr(pos+8,email,"#")-pos-8)
        call SendMail(who,email)
		
		ff2.write LifeAll & vbcrlf
		
		'if email<>"" then
		'	ff.write DT(now,"hh:mi:ss") & " " & who & " " & email & " (�ݧ�:" & ChkNum & "���A����:" & UpdNum & "��)" & vbcrlf
		'end if
      end if
  end if
  rs1.movenext
wend
rs1.close

'--------------------------------�{������------------------------
if Err.Number<>0 then ff.write vbcrlf & Err.Description & " " & vbcrlf
ff.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(End)-------------------------------------------------------------------" & vbcrlf & vbcrlf
ff.close

ff2.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(End)-------------------------------------------------------------------" & vbcrlf & vbcrlf
ff2.close

'msgbox "ok"
'--------------------------------���o�ͩR�i��----------------------------------------------------------------------------------------------------------
Sub GetLife(tbl,SQL)  
  if tbl<>"*" then
    SQL="select * from [�ͩR�i��] where [���ʤ��]>='" & DT(now-7,"yyyy/mm/dd 00:00:00") & "' and [���W��]='" & tbl _
      & "' and (" & SQL & ") order by [���W��],[���ʤH��],[�i���s��]"
  else
    SQL="select * from [�ͩR�i��] where [���ʤ��]>='" & DT(now-7,"yyyy/mm/dd 00:00:00") & "' and [���W��]<>'����]��' and [���W��]<>'�@�~�D��'" _
      & " and (" & SQL & ") order by [���W��],[���ʤH��],[�i���s��]"    
  end if
  'ff.write SQL & vbcrlf

  rs.open SQL,conn
  if not rs.eof then
    while not rs.eof
      LifeLog=rs("�ͩR�i��")	'�n��Ū�A�_�hŪ���X��
      if UpdNum<10 then LifeAll=LifeAll & rs("�i���s��") & "." & rs("���ʤH��") & "���� " & tbl & "(" & rs("�D��s��") & ")" & "�@�G " & vbcrlf _
        & LifeLog & " " & vbcrlf & " " & vbcrlf
      UpdNum=UpdNum+1
      rs.movenext
    wend
    flag=1
  end if
  rs.close
End Sub
'-------------------mail-----------------------------------------------------------------------------------------------------------------
Sub SendMail(who,email)	
  if email="" then exit sub
  ff.write DT(now,"hh:mi:ss") & " " & who & " " & email & " (�ݧ�:" & ChkNum & "���A����:" & UpdNum & "��)" & vbcrlf
 
  sch = "http://schemas.microsoft.com/cdo/configuration/" 
  Set cdoConfig=CreateObject("CDO.Configuration") 
  With cdoConfig.Fields 
	.Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
	.Item(sch & "smtpserverport") = 25
	.Item(sch & "smtpconnectiontimeout") = 60
	.Item(sch & "smtpserver") = "ms1.cwb.gov.tw"  'your smtp
	.update 
  End With
 
  Set cdoMessage=CreateObject("CDO.Message") 
  With cdoMessage 
	.Configuration = cdoConfig 
	.From     ="sosop@cwb.gov.tw"
	.To       =email
	.Subject  =DT(now," (mm/dd)") & " �귽�޲z�t��(IDMS)��Ƭd�֤β��ʳq��(�ݧ�:" & ChkNum & "���A����:" & UpdNum & "��)"
	' .HtmlBody =LifeAll
	.TextBody =LifeAll	'��r�Ҧ��~�L�ýX���D
	.BodyPart.charset = "UTF-8" 

	.Send
  End With 

  Set cdoConfig  =nothing
  Set cdoMessage =nothing
End Sub
'--------------------------------�ɶ��榡���----------------------------------------------------------------------------------------------------------
Function DT(byval sDT,byval fDT)
  DT=lcase(fDT)
  YY=year(sDT)   : if instr(1,DT,"yy",1)>0 and instr(1,DT,"yyyy",1)=0 then YY=mid(YY,3)
  DT=replace(replace(DT,"yyyy","yy"),"yy",YY)
  MM=month(sDT)  : if MM<10 then MM="0" & MM
  DT=replace(DT,"mm",MM)
  DD=day(sDT)    : if DD<10 then DD="0" & DD
  DT=replace(DT,"dd",DD)
  HH=hour(sDT)   : if HH<10 then HH="0" & HH
  DT=replace(DT,"hh",HH)
  MI=minute(sDT) : if MI<10 then MI="0" & MI
  DT=replace(DT,"mi",MI)
  SS=second(sDT) : if SS<10 then SS="0" & SS
  DT=replace(DT,"ss",SS) 
End Function
'--------------------------------���o�H��SQL���----------------------------------------------------------------------------------------------------------
Function GetSQL(fld,val)
  GetSQL = fld + "='" + val + "' or " + fld + " in (select Kind from Config where Item='" + val + "')" 
End Function
'--------------------------------��s��Ƭd�ָ��---------------------------------------------------------------------------------------------------------
Sub CheckDB()
  conn.execute "delete from [��Ƭd��]"
  rs.open "select * from [View_��Ƭd��] where [�d�ֵ��G]<>''",conn
  while not rs.eof
    conn.execute "insert into [��Ƭd��] values('" & rs(0) & "'," & rs(1) & ",'" & rs(2) & "','" & rs(3) & "','" & rs(4) & "','" & rs(5) & "')"
    rs.movenext
  wend
  rs.close
  ff.write DT(now,"hh:mi:ss") & " ������s[��Ƭd��]" & vbcrlf
End Sub
