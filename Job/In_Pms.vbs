'============================================================================================================================================================
'�\��G�C�ѶפJ�]���t�θ�Ʀ�IDMS�A�H�[�t�d�ߤαa�J�\��A�ûPISMS�P�B��s�겣�M�U
'============================================================================================================================================================
on error resume next		'��L�@��,���i�X��,�H�Khung��

'��Ʈw�s��--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
set fs=createobject("scripting.filesystemobject")
set ff=fs.OpenTextFile("In_Pms_" & DT(now,"yyyy") & ".log",8,true)
NowDT=DT(now,"yyyy/mm/dd hh:mi:ss")
ff.write NowDT & " "

strConn="Driver={SQL SERVER};Server=10.6.1.11;Trusted_Connection=True;Database=" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn & "IDMS" : conn.CommandTimeout = 1800
set rs=createobject("ADODB.Recordset")

LifeNo=1
rs.open "select max([�i���s��]) from [�ͩR�i��]",conn
if not rs.eof then LifeNo=rs(0)+1
rs.close  

'�PISMS�P�B��s�겣�M�U-----------------------------------------------------------------------------------------------------------------------
Pcount=0
rs.open "SELECT COUNT(*) FROM [View_�겣�M�U]",conn
Pcount=rs(0)
rs.close

if Pcount>10 and Err.Number=0 then  '�겣�M�U�s���W�h������
  'ISMS�s�W�겣�s��-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Value],[Description] from [View_�겣�M�U] where [Cancel]=0 and [Number] not in (select [Item] from [Config] where [Kind]='�겣�M�U')",conn
  while not rs.eof
    Memo=rs(3)	'���o�˨��ȡA�|�ܪť�
    conn.execute "insert into [Config] values('�겣�M�U','" & rs(0) & "','" & rs(1) & "','" & rs(2) & "','" & Memo & "')"
    conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'�겣�M�U',0,'�s�W�G" & rs(0) & " " & rs(1) & " " & Memo & "','SSM�p��','','','����OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close
  'ISMS�ק�겣�W�٩δy�z-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Description],[Config],[Memo],[Value],[Mark] from [View_�겣�M�U],[Config] where [View_�겣�M�U].[Number]=[Config].[Item] and [Cancel]=0 and [Kind]='�겣�M�U' and ([Name]<>[Config] or [Description]<>[Memo])",conn
  while not rs.eof  '�겣���Ȫ����ʧP�_�]�bMemo���A���t�~�B�z
    Memo=rs(2)	'���o�˨��ȡA�|�ܪť�
    Life="�ק�G" & rs(0) & " " & rs(3) & "->" & rs(1) & " " & rs(4) & "->" & Memo
    conn.execute "Update [Config] set [Config]='" & rs(1) & "',[Mark]='" & rs(5) & "',[Memo]='" & Memo & "' where [Kind]='�겣�M�U' and [Item]='" & rs(0) & "'"
    conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'�겣�M�U',0,'�ק�G" & rs(0) & " " & rs(3) & "->" & rs(1) & " " & rs(4) & "->" & Memo & "','SSM�p��','','','����OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close

  'ISMS�R���겣�s��-----------------------------------------------------------------------------------------------------------------------
  rs.open "select * from [Config] where [Kind]='�겣�M�U' and [Config] not like '(�w���P)%' and [Item] not in (select [Number] from [View_�겣�M�U] where [Cancel]=0)",conn,3,3
  while not rs.eof    
    conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'�겣�M�U',0,'�R���G" & rs(0) & " " & rs(1) & " " & rs(2) & " (�Щ�IDMS�t�γ]�w��ʧR���A�ê`�N��Ƭ̩ۨ�)','SSM�p��','','','����OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    rs(2)="(�w���P)" & rs(2)
    rs.update
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close

  ff.write "�겣�M�U�G" & DT(now,"hh:mi:ss") & " "
end if

'�פJ���`�]��-----------------------------------------------------------------------------------------------------------------------
conn.execute "drop table [���`�]��]"
conn.execute "select * into [���`�]��] from [View_���`�]��]"
ff.write "���`�]���G" & DT(now,"hh:mi:ss") & " "    

'���`�]�����ʬd��-----------------------------------------------------------------------------------------------------------------------
if Err.Number=0 then    '���`�]���s�������\�h������
  Pcount=0
  rs.open "SELECT COUNT(*) FROM [���`�]��]",conn
  Pcount=rs(0)
  rs.close

  if Pcount>5000 then  '���`�]���פJ�����\�h������
    '�ˬd���`�]���O�ޤH�ܧ�-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [�]�ƽs��],[�]���D��].[�]���s��A] as [�]���s��A],[�]���D��].[�]���s��B] as [�]���s��B],[�]���D��].[�]���O�W] as [�]���O�W]" _
      & ",[���@�H��],[���`�]��].[�O�ޤH��] as [�O�ޤH��],[�]���D��].[�O�ޤH��] as [��O�ޤH]" _
      & " from [�]���D��],[���`�]��],[����]��]" _
      & " where [�]���D��].[�]���s��A]=[���`�]��].[�]���s��A] and [�]���D��].[�]���s��B]=[���`�]��].[�]���s��B]" _
      & " and   [�]���D��].[�]���s��A]=[����]��].[�]���s��A] and [�]���D��].[�]���s��B]=[����]��].[�]���s��B]" _
      & " and [�]���D��].[�O�ޤH��]<>[���`�]��].[�O�ޤH��]",conn
    while not rs.eof 
      conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'���`�]��'," & rs(0) _
        & ",'���`�]���O�ޤH�ܧ�(" & rs(1) & " " & rs(2) & " " & rs(3) & ")�G" & rs(6) & "->" & rs(5) & "'" _
        & ",'" & rs(4) & "','�]���q��','" & rs(6) & "','�]�ޤH��','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '�ˬd���`�]�����o-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [�]���D��].[�]���s��A] + ' ' + [�]���D��].[�]���s��B] as [�]���s��],[�]���D��].[�]���O�W]" _
      & ",[�]���D��].[�t�P],[�]���D��].[����],[�]���D��].[�ʶR���],[�]���D��].[�O�ޤH��]" _
      & " from [�]���D��],[���`�]��] where [�]���D��].[�]���s��A]=[���`�]��].[�]���s��A] and [�]���D��].[�]���s��B]=[���`�]��].[�]���s��B]" _
      & " and [�]���D��].[�L�ĵ��O]='' and [���`�]��].[�L�ĵ��O]='N'",conn    
    while not rs.eof 
      conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'���`�]��',0" _
        & ",'���`�]�����o" & "(" & rs(0) & " " & rs(1) & ")�G" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','�]���q��','','�]�ޤH��','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '�ˬd���`�]���s�W-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [�]���s��A] + ' ' + [�]���s��B] as [�]���s��],[�]���O�W],[�t�P],[����],[�ʶR���],[�O�ޤH��] from [���`�]��]" _
      & " where [�]���s��A] + ' ' + [�]���s��B] not in (select [�]���s��A] + ' ' + [�]���s��B] from [�]���D��])",conn
    while not rs.eof     
      conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'���`�]��',0" _
        & ",'���`�]���s�W(" & rs(0) & " " & rs(1) & ")�G" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','�]���q��','','�]�ޤH��','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '�ˬd���`�]������-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [�]���s��A] + ' ' + [�]���s��B] as [�]���s��],[�]���O�W],[�t�P],[����],[�ʶR���],[�O�ޤH��] from [�]���D��]" _
      & " where [�]���s��A] + ' ' + [�]���s��B] not in (select [�]���s��A] + ' ' + [�]���s��B] from [���`�]��])",conn
    while not rs.eof 
      conn.execute "insert into [�ͩR�i��] values(" & LifeNo & ",'���`�]��',0" _
        & ",'���`�]������(" & rs(0) & " " & rs(1) & ")�G" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','�]���q��','','�]�ޤH��','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    ff.write "���ʬd�֡G" & DT(now,"hh:mi:ss") & " " 

    '��s�]���D��-----------------------------------------------------------------------------------------------------------------------
    conn.execute "drop table [�]���D��]"
    conn.execute "select * into [�]���D��] from [���`�]��]" 
    conn.execute "ALTER TABLE [�]���D��] ADD  CONSTRAINT [PK_�]���D��] PRIMARY KEY CLUSTERED ([�����N��] ASC,[�]���s��A] ASC,[�]���s��B] ASC)" _
        & " WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,SORT_IN_TEMPDB=OFF,IGNORE_DUP_KEY=OFF,ONLINE=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    conn.execute "UPDATE [�]���D��] SET [�]���O�W]='' where [�]���O�W] is null"
    conn.execute "UPDATE [�]���D��] SET [�t�P]='' where [�t�P] is null"
    conn.execute "UPDATE [�]���D��] SET [����]='' where [����] is null"
    conn.execute "UPDATE [�]���D��] SET [�p�q���]='' where [�p�q���] is null"
    conn.execute "UPDATE [�]���D��] SET [�ϥΦ~��]=0 where [�ϥΦ~��] is null"
    conn.execute "UPDATE [�]���D��] SET [���u�N��]='' where [���u�N��] is null"
    conn.execute "UPDATE [�]���D��] SET [�]���a�I]='' where [�]���a�I] is null"
    conn.execute "UPDATE [�]���D��] SET [�O�޽ҧO]='' where [�O�޽ҧO] is null"
    conn.execute "UPDATE [�]���D��] SET [�O�ޤH��]='' where [�O�ޤH��] is null"
    conn.execute "UPDATE [�]���D��] SET [�L�ĵ��O]='' where [�L�ĵ��O] is null"
    conn.execute "UPDATE [�]���D��] SET [�]������]='' where [�]������] is null"

    ff.write "�]���D�ɡG" & DT(now,"hh:mi:ss") & " " 

  end if  '���`�]���פJ�����\�h������
end if    '���`�]���s�������\�h������(Err.Number)
'------------------------------------------------------------
if Err.Number<>0 then ff.write Err.Description & " "

ff.write DT(now,"hh:mi:ss") & " (" & Pcount & ")" & vbcrlf
ff.close

'msgbox "ok!"
'--------------------------------�ɶ��榡���---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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