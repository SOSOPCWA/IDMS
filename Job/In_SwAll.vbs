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

Set reNum=new RegExp   '數字
reNum.IgnoreCase =true
reNum.Global=True
reNum.Pattern="^[+-]?\d+$"
'-----------------------------------------------------------------啟始處理
NowDT=DT(now,"yyyy/mm/dd hh:mi")
dim SwNo   : SwNo=0
dim sidnum : sidnum="0"
dim AskNo  : AskNo=1
conn.execute "delete from [軟體主檔]"
conn.execute "delete from [軟體授權]"
conn.execute "delete from [申請表單]"
'conn.execute "delete from [生命履歷] where [表格名稱]='軟體主檔' or [表格名稱]='軟體授權'"
'--------------------------------------------------------------------------------------------------------------匯入[軟體主檔]
rs.open "select * from [軟體管理] where [保管單編號] is not null order by [保管單編號] DESC,[sidnum]",conn
while not rs.eof  
    SwFirstYN=""
    if rs("保管單編號")<>SwFormNo then 
        SwFirstYN="Y"    '是否是第一筆資料，若是，需匯入[軟體主檔]  
        sidnum="0"    'PK=SwNo+sidnum，故歸零之   
        ff.write vbcrlf '換保管單即多換一行
        SwNo=SwNo+1
    end if

    SwFormNo=rs("保管單編號")
    'Wscript.Echo SwFormNo
    'SwFormNo.Print
    if len(SwFormNo)<>9 then
        fe.write SwNo & "(軟體) " & SwFormNo & " " & rs("sidnum") & " 表單編號：" & SwFormNo & vbcrlf
        if len(SwFormNo)>9 then SwFormNo=mid(SwFormNo,1,9)   '避免輸入錯誤資料
    end if

    SwUnit=""   '軟體申請單位
    SwAsker=""  '軟體申請人員
    SwName="" : if not IsNull(rs("軟體名稱")) then SwName=trim(rs("軟體名稱"))  '軟體名稱
    VersBuy="" : if not IsNull(rs("購買版本")) then VersBuy=replace(trim(rs("購買版本")),"'","")  '購買版本
    VersCan="" : if not IsNull(rs("可用版本")) then VersCan=trim(rs("可用版本"))  '可用版本

    LcsKind="" : if not IsNull(rs("授權方式")) then LcsKind=trim(rs("授權方式"))  '授權方式
    LcsNum="0" : if not IsNull(rs("授權數量")) then LcsNum=trim(rs("授權數量"))  '授權數量
    SwLcsMemo="" '授權說明

    Source="" : if not IsNull(rs("軟體來源")) then Source=trim(rs("軟體來源"))  '軟體來源
    Functions="" : if not IsNull(rs("軟體功用")) then Functions=replace(trim(rs("軟體功用")),"'","")  '軟體功能

    SwBrand="" : if not IsNull(rs("適用廠牌")) then SwBrand=trim(rs("適用廠牌"))  '授權廠牌【都匯】
    SwStyle="" : if not IsNull(rs("適用型式")) then SwStyle=trim(rs("適用型式"))  '授權型式【都匯】

    SnBuy=""  '購買序號
    SnDown="" '降級序號

    Life="" : if not IsNull(rs("期限說明")) then Life=trim(rs("期限說明"))  '期限說明

    BuyDay="null"
    if not IsNull(rs("購買日期")) then 
      BuyDay="'" & trim(rs("購買日期")) & "'"  '購買日期
      if not reDate.Test(BuyDay) then
        fe.write SwNo & "(軟體) " & SwFormNo & " " & rs("sidnum") & " 購買日期：" & BuyDay & vbcrlf
        BuyDay="null"   '避免輸入錯誤資料
      end if
    end if

    UseDay="null"
    if not IsNull(rs("使用期限")) then 
      UseDay="'" & trim(rs("使用期限")) & "'"  '使用期限
      if UseDay="'永久'" then 
        UseDay="'2099/12/31'"
        if Life="" then 
            Life="永久"
        else
            Life=Life & "，永久"
        end if
      elseif not reDate.Test(UseDay) then
        fe.write SwNo & "(軟體) " & SwFormNo & " " & rs("sidnum") & " 使用期限：" & UseDay & vbcrlf
        UseDay="null"   '避免輸入錯誤資料
      end if
    end if

    UpDay="null" '更新期限
    if not IsNull(rs("更新日期")) then 
      UpDay="'" & trim(rs("更新日期")) & "'"
      if UpDay="'啟用日一年內'" then 
        UpDay="null"
        if Life="" then 
            Life="序號啟用後1年"
        else
            Life=Life & "，序號啟用後1年"
        end if
      elseif not reDate.Test(UpDay) then
        fe.write SwNo & "(軟體) " & SwFormNo & " " & rs("sidnum") & " 更新期限：" & UpDay & vbcrlf
        UpDay="null"   '避免輸入錯誤資料
      end if
    end if

    Price=0 : if not IsNull(rs("軟體單價")) then Price=trim(rs("軟體單價"))  '軟體單價
    Total=0 : if not IsNull(rs("軟體總價")) then Total=trim(rs("軟體總價"))  '軟體總價
    PriceMemo="" '價格說明

    Supply="" : if not IsNull(rs("提供單位")) then Supply=trim(rs("提供單位"))  '提供單位
    Media="" '存放媒體

    BookNo="" : if not IsNull(rs("圖書編號")) then BookNo=trim(rs("圖書編號"))  '圖書編號
    SwAttach="" : if not IsNull(rs("軟體附件")) then SwAttach=trim(rs("軟體附件"))  '軟體附件【都匯】

    SwStatus="使用中" : if not IsNull(rs("軟體狀態")) then SwStatus=trim(rs("軟體狀態")) '軟體狀態
    if SwStatus="" then SwStatus="可使用"

    SwCause="" '減損原因
    SwExec="" '減損方式
    SwStatusDay="null" '減損日期
    SwKeyDay="null" '填表日期

    SwCreateDay="null" '建立日期
    SwCreater="" '建立人員
    SwSaveDay="null" '修改日期
    SwSaver="" '修改人員
    SwMemo="" : if not IsNull(rs("備註一")) then SwMemo=trim(replace(rs("備註一"),"'",""))  '備註說明【都匯】

    SwSQL="insert into [軟體主檔] values(" _
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

    if SwFirstYN="Y" then   '是否是第一筆資料，若是，需匯入[軟體主檔]
        ff.write SwSQL & vbcrlf 
        conn.execute SwSQL        
    end if
    '--------------------------------------------------------------------------------------------------------------匯入[軟體授權]
    ApNo=0
    if not IsNull(rs("作業編號")) then 
        ApNo=trim(rs("作業編號")) : if ApNo="?" then ApNo=0  '作業編號
        if not reNum.Test(ApNo) then
            fe.write AskNo & "(授權) " & SwFormNo & " " & rs("sidnum") & " 作業編號：" & ApNo & vbcrlf
            ApNo=0   '避免輸入錯誤資料
        end if
    end if

    sidnum="0" : if not IsNull(rs("sidnum")) then sidnum=trim(rs("sidnum"))  '授權識別
    if sidnum="" or sidnum="0" then  '授權識別碼重複或空白
        fe.write AskNo & "(授權) " & SwFormNo & " " & rs("sidnum") & " 授權識別值空白" & vbcrlf
    else
        if not reNum.Test(sidnum) then fe.write AskNo & "(授權) " & SwFormNo & " " & rs("sidnum") & " 授權識別值有誤" & vbcrlf    
    end if

    VersAsk="" '申請版本
    LcsAsk="" '申請授權
    LcsMemoAsk="" '授權說明
    SnAsk="" '申請序號

    UnitLcs="" : if not IsNull(rs("申請單位")) then UnitLcs=trim(rs("申請單位"))  '授權單位
    UserLcs="" : if not IsNull(rs("申請人員")) then UserLcs=trim(rs("申請人員"))  '授權人員
    HostLcs="" : if not IsNull(rs("申請主機")) then HostLcs=trim(rs("申請主機"))  '授權主機
    IpLcs="" : if not IsNull(rs("申請IP")) then IpLcs=trim(rs("申請IP"))  '授權IP
    PropNoALcs="" : if not IsNull(rs("申請財編A")) then PropNoALcs=trim(rs("申請財編A"))  '授權財編A
    PropNoBLcs="" : if not IsNull(rs("申請財編B")) then PropNoBLcs=trim(rs("申請財編B"))  '授權財編B
    BrandLcs=SwBrand   '授權廠牌【都匯】
    StyleLcs=SwStyle   '授權型式【都匯】
    
    AskAttach=SwAttach  '授權附件【都匯】

    AskStatus="已申請" : if not IsNull(rs("授權狀態")) then AskStatus=trim(rs("授權狀態")) '授權狀態
    if AskStatus="" then AskStatus="已申請"

    AskCause="" '減損原因
    AskExec="" '減損方式
    AskStatusDay="null" '減損日期

    AskKeyDay="null" '填表日期******************【匯領用日期】
    if not IsNull(rs("領用日期")) then 
      AskKeyDay="'" & trim(rs("領用日期")) & "'" 
      if AskKeyDay="'0'" then 
        AskKeyDay="null" 
      elseif not reDate.Test(AskKeyDay) then
        fe.write AskNo & "(授權) " & SwFormNo & " " & rs("sidnum") & " 領用日期：" & AskKeyDay & vbcrlf
        AskKeyDay="null"   '避免輸入錯誤資料
      end if
    end if

    AskKeyer="" '填表人員

    AskCreateDay="null" '建立日期
    AskCreater="" '建立人員
    AskSaveDay="null" '修改日期
    AskSaver="" '修改人員
    AskMemo="" : if not IsNull(rs("備註一")) then SwMemo=trim(replace(rs("備註一"),"'",""))  '備註說明【都匯】

    AskSQL="insert into [軟體授權] values(" _
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
    conn.execute AskSQL '匯入[軟體授權]
    '--------------------------------------------------------------------------------------------------------------匯入[申請表單]
    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("申請事項")) then AskItem=trim(rs("申請事項"))
    if not IsNull(rs("表單號碼")) then AskFormNo=trim(rs("表單號碼"))
    if not IsNull(rs("填表日期")) then AskKeyDate="'" & trim(rs("填表日期")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"0")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("申請事項1")) then AskItem=trim(rs("申請事項1"))
    if not IsNull(rs("表單號碼1")) then AskFormNo=trim(rs("表單號碼1"))
    if not IsNull(rs("填表日期1")) then AskKeyDate="'" & trim(rs("填表日期1")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"1")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("申請事項2")) then AskItem=trim(rs("申請事項2"))
    if not IsNull(rs("表單號碼2")) then AskFormNo=trim(rs("表單號碼2"))
    if not IsNull(rs("填表日期2")) then AskKeyDate="'" & trim(rs("填表日期2")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"2")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("申請事項3")) then AskItem=trim(rs("申請事項3"))
    if not IsNull(rs("表單號碼3")) then AskFormNo=trim(rs("表單號碼3"))
    if not IsNull(rs("填表日期3")) then AskKeyDate="'" & trim(rs("填表日期3")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"3")

    AskItem="" : AskFormNo="" : AskKeyDate="''"
    if not IsNull(rs("申請事項4")) then AskItem=trim(rs("申請事項4"))
    if not IsNull(rs("表單號碼4")) then AskFormNo=trim(rs("表單號碼4"))
    if not IsNull(rs("填表日期4")) then AskKeyDate="'" & trim(rs("填表日期4")) & "'"
    call InsAskForm(AskNo,AskItem,AskFormNo,AskKeyDate,"4")    
    '--------------------------------------------------------------------------------------------------------------
    Wscript.Echo AskNo
	AskNo=AskNo+1   
  rs.movenext
wend
rs.close

'--------------------------------------------------------------------------------------------------------------結束處理
ff.close
fe.close
msgbox "OK!"
'--------------------------------時間格式函數-------------------------------------
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
        if AskFormNo="" then AskFormNo="(空白)"
        if AskItem="" then AskItem="(空白)"             
            
        if AskKeyDate="''" then 
            AskKeyDate="null"
        else
            if mid(AskKeyDate,2,1)="1" then AskKeyDate="'" & Cstr(Cint(mid(AskKeyDate,2,3))+1911) & mid(AskKeyDate,5) '將民國年改為西元年

            if not reDate.Test(AskKeyDate) then
                fe.write AskNo & "(表單) " & SwFormNo & " " & rs("sidnum") & " 填表日期：" & AskItem & " " & AskKeyDate & vbcrlf
                AskKeyDate="null"   '避免輸入錯誤資料
            end if
        end if

        FormSQL="insert into [申請表單] values('軟體申請'," & AskNo & ",'" & AskFormNo & "-" & i & "','" & AskItem & "'," & AskKeyDate & ")"
        ff.write FormSQL & "：" & SwFormNo & vbcrlf
        conn.execute FormSQL
    end if
End Sub