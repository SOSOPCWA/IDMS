'on error resume next
'-----------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=10.6.1.4;Trusted_Connection=True;Database=IDMS" 
'strConn="Driver={SQL SERVER};Server=(local)\SQLEXPRESS;Trusted_Connection=True;Database=IDMS" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn
Dim rs : set rs=createobject("ADODB.Recordset")

set fs=createobject("scripting.filesystemobject")
set fe=fs.OpenTextFile("In_Error.log",2,true)
set ff=fs.OpenTextFile("In_SwAll.log",2,true)

Set reDate=new RegExp   '['YYYY/MM/DD']
reDate.IgnoreCase =true
reDate.Global=True
reDate.Pattern="^['](((19|20)[0-9]{2})[- /.](0?[1-9]|1[012])[- /.](0?[1-9]|[12][0-9]|3[01]))[']$"

Set reNum=new RegExp   '�Ʀr
reNum.IgnoreCase =true
reNum.Global=True
reNum.Pattern="^[+-]?\d+$"
'-----------------------------------------------------------------�ҩl�B�z
NowDT=DT(now,"yyyy/mm/dd hh:mi")
dim SwNo   : SwNo=0
dim sidnum : sidnum="0"
dim AskNo  : AskNo=1
conn.execute "delete from [�n��D��]"
conn.execute "delete from [�n����v]"
conn.execute "delete from [�ӽЪ��]"
'conn.execute "delete from [�ͩR�i��] where [���W��]='�n��D��' or [���W��]='�n����v'"
'--------------------------------------------------------------------------------------------------------------�פJ[�n��D��]
rs.open "select * from [�n��޲z] where [�O�޳�s��] is not null order by [�O�޳�s��] DESC,[sidnum]",conn
while not rs.eof  
    SwFirstYN=""
    if rs("�O�޳�s��")<>SwFormNo then 
        SwFirstYN="Y"    '�O�_�O�Ĥ@����ơA�Y�O�A�ݶפJ[�n��D��]  
        sidnum="0"    'PK=SwNo+sidnum�A�G�k�s��   
        ff.write vbcrlf '���O�޳�Y�h���@��
        SwNo=SwNo+1
    end if

    SwFormNo=rs("�O�޳�s��")
    'Wscript.Echo SwFormNo
    'SwFormNo.Print
    if len(SwFormNo)<>9 then
        fe.write SwNo & "(�n��) " & SwFormNo & " " & rs("sidnum") & " ���s���G" & SwFormNo & vbcrlf
        if len(SwFormNo)>9 then SwFormNo=mid(SwFormNo,1,9)   '�קK��J���~���
    end if

    SwUnit=""   '�n��ӽг��
    SwAsker=""  '�n��ӽФH��
    SwName="" : if not IsNull(rs("�n��W��")) then SwName=trim(rs("�n��W��"))  '�n��W��
    VersBuy="" : if not IsNull(rs("�ʶR����")) then VersBuy=replace(trim(rs("�ʶR����")),"'","")  '�ʶR����
    VersCan="" : if not IsNull(rs("�i�Ϊ���")) then VersCan=trim(rs("�i�Ϊ���"))  '�i�Ϊ���

    LcsKind="" : if not IsNull(rs("���v�覡")) then LcsKind=trim(rs("���v�覡"))  '���v�覡
    LcsNum="0" : if not IsNull(rs("���v�ƶq")) then LcsNum=trim(rs("���v�ƶq"))  '���v�ƶq
    SwLcsMemo="" '���v����

    Source="" : if not IsNull(rs("�n��ӷ�")) then Source=trim(rs("�n��ӷ�"))  '�n��ӷ�
    Functions="" : if not IsNull(rs("�n��\��")) then Functions=replace(trim(rs("�n��\��")),"'","")  '�n��\��

    SwBrand="" : if not IsNull(rs("�A�μt�P")) then SwBrand=trim(rs("�A�μt�P"))  '���v�t�P�i���סj
    SwStyle="" : if not IsNull(rs("�A�Ϋ���")) then SwStyle=trim(rs("�A�Ϋ���"))  '���v�����i���סj

    SnBuy=""  '�ʶR�Ǹ�
    SnDown="" '���ŧǸ�

    Life="" : if not IsNull(rs("��������")) then Life=trim(rs("��������"))  '��������

    BuyDay="null"
    if not IsNull(rs("�ʶR���")) then 
      BuyDay="'" & trim(rs("�ʶR���")) & "'"  '�ʶR���
      if not reDate.Test(BuyDay) then
        fe.write SwNo & "(�n��) " & SwFormNo & " " & rs("sidnum") & " �ʶR����G" & BuyDay & vbcrlf
        BuyDay="null"   '�קK��J���~���
      end if
    end if

    UseDay="null"
    if not IsNull(rs("�ϥδ���")) then 
      UseDay="'" & trim(rs("�ϥδ���")) & "'"  '�ϥδ���
      if UseDay="'�ä['" then 
        UseDay="'2099/12/31'"
        if Life="" then 
            Life="�ä["
        else
            Life=Life & "�A�ä["
        end if
      elseif not reDate.Test(UseDay) then
        fe.write SwNo & "(�n��) " & SwFormNo & " " & rs("sidnum") & " �ϥδ����G" & UseDay & vbcrlf
        UseDay="null"   '�קK��J���~���
      end if
    end if

    UpDay="null" '��s����
    if not IsNull(rs("��s���")) then 
      UpDay="'" & trim(rs("��s���")) & "'"
      if UpDay="'�ҥΤ�@�~��'" then 
        UpDay="null"
        if Life="" then 
            Life="�Ǹ��ҥΫ�1�~"
        else
            Life=Life & "�A�Ǹ��ҥΫ�1�~"
        end if
      elseif not reDate.Test(UpDay) then
        fe.write SwNo & "(�n��) " & SwFormNo & " " & rs("sidnum") & " ��s�����G" & UpDay & vbcrlf
        UpDay="null"   '�קK��J���~���
      end if
    end if

    Price=0 : if not IsNull(rs("�n����")) then Price=trim(rs("�n����"))  '�n����
    Total=0 : if not IsNull(rs("�n���`��")) then Total=trim(rs("�n���`��"))  '�n���`��
    PriceMemo="" '���满��

    Supply="" : if not IsNull(rs("���ѳ��")) then Supply=trim(rs("���ѳ��"))  '���ѳ��
    Media="" '�s��C��

    BookNo="" : if not IsNull(rs("�Ϯѽs��")) then BookNo=trim(rs("�Ϯѽs��"))  '�Ϯѽs��
    SwAttach="" : if not IsNull(rs("�n�����")) then SwAttach=trim(rs("�n�����"))  '�n�����i���סj

    SwStatus="�ϥΤ�" : if not IsNull(rs("�n�骬�A")) then SwStatus=trim(rs("�n�骬�A")) '�n�骬�A
    if SwStatus="" then SwStatus="�i�ϥ�"

    SwCause="" '��l��]
    SwExec="" '��l�覡
    SwStatusDay="null" '��l���
    SwKeyDay="null" '�����

    SwCreateDay="null" '�إߤ��
    SwCreater="" '�إߤH��
    SwSaveDay="null" '�ק���
    SwSaver="" '�ק�H��
    SwMemo="" : if not IsNull(rs("�Ƶ��@")) then SwMemo=trim(replace(rs("�Ƶ��@"),"'",""))  '�Ƶ������i���סj

    SwSQL="insert into [�n��D��] values(" _
      & SwNo & ",'" _
      & SwFormNo & "','" _
      & SwUnit & "','" _
      & SwAsker & "','" _
      & SwName & "','" _
      & VersBuy & "','" _
      & VersCan & "','" _
      & LcsKind & "'," _
      & LcsNum & ",'" _
      & SwLcsMemo & "','" _
      & Source & "','" _
      & Functions & "','" _
      & SwBrand & "','" _
      & SwStyle & "','" _
      & SnBuy & "','" _
      & SnDown & "','" _
      & Life & "'," _
      & BuyDay & "," _
      & UseDay & "," _
      & UpDay & "," _
      & Price & "," _
      & Total & ",'" _
      & PriceMemo & "','" _
      & Supply & "','" _
      & Media & "','" _
      & BookNo & "','" _
      & SwAttach & "','" _
      & SwStatus & "','" _
      & SwCause & "','" _
      & SwExec & "'," _
      & SwStatusDay & "," _
      & SwKeyDay & "," _
      & SwCreateDay & ",'" _
      & SwCreater & "'," _
      & SwSaveDay & ",'" _
      & SwSaver & "','" _
      & SwMemo & "')"

    if SwFirstYN="Y" then   '�O�_�O�Ĥ@����ơA�Y�O�A�ݶפJ[�n��D��]
        ff.write SwSQL & vbcrlf 
        conn.execute SwSQL        
    end if
    '--------------------------------------------------------------------------------------------------------------�פJ[�n����v]
    ApNo=0
    if not IsNull(rs("�@�~�s��")) then 
        ApNo=trim(rs("�@�~�s��")) : if ApNo="?" then ApNo=0  '�@�~�s��
        if not reNum.Test(ApNo) then
            fe.write AskNo & "(���v) " & SwFormNo & " " & rs("sidnum") & " �@�~�s���G" & ApNo & vbcrlf
            ApNo=0   '�קK��J���~���
        end if
    end if

    sidnum="0" : if not IsNull(rs("sidnum")) then sidnum=trim(rs("sidnum"))  '���v�ѧO
    if sidnum="" or sidnum="0" then  '���v�ѧO�X���ƩΪť�
        fe.write AskNo & "(���v) " & SwFormNo & " " & rs("sidnum") & " ���v�ѧO�Ȫť�" & vbcrlf
    else
        if not reNum.Test(sidnum) then fe.write AskNo & "(���v) " & SwFormNo & " " & rs("sidnum") & " ���v�ѧO�Ȧ��~" & vbcrlf    
    end if

    VersAsk="" '�ӽЪ���
    LcsAsk="" '�ӽб��v
    LcsMemoAsk="" '���v����
    SnAsk="" '�ӽЧǸ�

    UnitLcs="" : if not IsNull(rs("�ӽг��")) then UnitLcs=trim(rs("�ӽг��"))  '���v���
    UserLcs="" : if not IsNull(rs("�ӽФH��")) then UserLcs=trim(rs("�ӽФH��"))  '���v�H��
    HostLcs="" : if not IsNull(rs("�ӽХD��")) then HostLcs=trim(rs("�ӽХD��"))  '���v�D��
    IpLcs="" : if not IsNull(rs("�ӽ�IP")) then IpLcs=trim(rs("�ӽ�IP"))  '���vIP
    PropNoALcs="" : if not IsNull(rs("�ӽа]�sA")) then PropNoALcs=trim(rs("�ӽа]�sA"))  '���v�]�sA
    PropNoBLcs="" : if not IsNull(rs("�ӽа]�sB")) then PropNoBLcs=trim(rs("�ӽа]�sB"))  '���v�]�sB
    BrandLcs=SwBrand   '���v�t�P�i���סj
    StyleLcs=SwStyle   '���v�����i���סj
    
    AskAttach=SwAttach  '���v����i���סj

    AskStatus="�w�ӽ�" : if not IsNull(rs("���v���A")) then AskStatus=trim(rs("���v���A")) '���v���A
    if AskStatus="" then AskStatus="�w�ӽ�"

    AskCause="" '��l��]
    AskExec="" '��l�覡
    AskStatusDay="null" '��l���

    AskKeyDay="null" '�����******************�i�׻�Τ���j
    if not IsNull(rs("��Τ��")) then 
      AskKeyDay="'" & trim(rs("��Τ��")) & "'" 
      if AskKeyDay="'0'" then 
        AskKeyDay="null" 
      elseif not reDate.Test(AskKeyDay) then
        fe.write AskNo & "(���v) " & SwFormNo & " " & rs("sidnum") & " ��Τ���G" & AskKeyDay & vbcrlf
        AskKeyDay="null"   '�קK��J���~���
      end if
    end if

    AskKeyer="" '���H��

    AskCreateDay="null" '�إߤ��
    AskCreater="" '�إߤH��
    AskSaveDay="null" '�ק���
    AskSaver="" '�ק�H��
    AskMemo="" : if not IsNull(rs("�Ƶ��@")) then SwMemo=trim(replace(rs("�Ƶ��@"),"'",""))  '�Ƶ������i���סj

    AskSQL="insert into [�n����v] values(" _
      & AskNo & "," _
      & SwNo & "," _
      & ApNo & "," _
      & sidnum & ",'" _
      & VersAsk & "','" _
      & LcsAsk & "','" _
      & LcsMemoAsk & "','" _
      & SnAsk & "','" _
      & UnitLcs & "','" _
      & UserLcs & "','" _
      & HostLcs & "','" _
      & IpLcs & "','" _
      & PropNoALcs & "','" _
      & PropNoBLcs & "','" _
      & BrandLcs & "','" _
      & StyleLcs & "','" _
      & AskAttach & "','" _
      & AskStatus & "','" _
      & AskCause & "','" _
      & VAskExec & "'," _
      & AskStatusDay & "," _
      & AskKeyDay & ",'" _
      & AskKeyer & "'," _
      & AskCreateDay & ",'" _
      & AskCreater & "'," _
      & AskSaveDay & ",'" _
      & AskSaver & "','" _
      & AskMemo & "')"

    'ff.write AskSQL & vbcrlf
    conn.execute AskSQL '�פJ[�n����v]
    '--------------------------------------------------------------------------------------------------------------�פJ[�ӽЪ��]
    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("�ӽШƶ�")) then AskItem=trim(rs("�ӽШƶ�"))
    if not IsNull(rs("��渹�X")) then AskFormNo=trim(rs("��渹�X"))
    if not IsNull(rs("�����")) then AskKeyDate="'" & trim(rs("�����")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"0")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("�ӽШƶ�1")) then AskItem=trim(rs("�ӽШƶ�1"))
    if not IsNull(rs("��渹�X1")) then AskFormNo=trim(rs("��渹�X1"))
    if not IsNull(rs("�����1")) then AskKeyDate="'" & trim(rs("�����1")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"1")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("�ӽШƶ�2")) then AskItem=trim(rs("�ӽШƶ�2"))
    if not IsNull(rs("��渹�X2")) then AskFormNo=trim(rs("��渹�X2"))
    if not IsNull(rs("�����2")) then AskKeyDate="'" & trim(rs("�����2")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"2")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("�ӽШƶ�3")) then AskItem=trim(rs("�ӽШƶ�3"))
    if not IsNull(rs("��渹�X3")) then AskFormNo=trim(rs("��渹�X3"))
    if not IsNull(rs("�����3")) then AskKeyDate="'" & trim(rs("�����3")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"3")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("�ӽШƶ�4")) then AskItem=trim(rs("�ӽШƶ�4"))
    if not IsNull(rs("��渹�X4")) then AskFormNo=trim(rs("��渹�X4"))
    if not IsNull(rs("�����4")) then AskKeyDate="'" & trim(rs("�����4")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"4")    
    '--------------------------------------------------------------------------------------------------------------
    Wscript.Echo AskNo
	AskNo=AskNo+1   
  rs.movenext
wend
rs.close

'--------------------------------------------------------------------------------------------------------------�����B�z
ff.close
fe.close
msgbox "OK!"
'--------------------------------�ɶ��榡���-------------------------------------
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

Sub InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,i)
    if AskItem<>"" or AskFormNo<>"" or AskKeyDate<>"''" then
        if AskFormNo="" then AskFormNo="(�ť�)"
        if AskItem="" then AskItem="(�ť�)"             
            
        if AskKeyDate="''" then 
            AskKeyDate="null"
        else
            if mid(AskKeyDate,2,1)="1" then AskKeyDate="'" & Cstr(Cint(mid(AskKeyDate,2,3))+1911) & mid(AskKeyDate,5) '�N����~�אּ�褸�~

            if not reDate.Test(AskKeyDate) then
                fe.write AskNo & "(���) " & SwFormNo & " " & rs("sidnum") & " ������G" & AskItem & " " & AskKeyDate & vbcrlf
                AskKeyDate="null"   '�קK��J���~���
            end if
        end if

        FormSQL="insert into [�ӽЪ��] values('�n��ӽ�'," & AskNo & ",'" & AskFormNo & "-" & i & "','" & AskItem & "'," & AskKeyDate & ")"
        ff.write FormSQL & "�G" & SwFormNo & vbcrlf
        conn.execute FormSQL
    end if
End Sub