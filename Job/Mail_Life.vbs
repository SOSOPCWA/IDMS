'============================================================================================================================================
'功能：每週一自動mail各維護人員他應該知道的事，主要是主檔的異動資訊(生命履歷)
'============================================================================================================================================
on error resume next		'寧無作用,不可出錯,以免hung住
MailTag="#{mail}:"
LeafTag="#{已離職}："
' 每行不加 & vbcrlf 會出錯
'--------------------------------資料庫連結--------------------------------------------------------------------------------------------------
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


call CheckDB()	'更新資料查核資料

ChkMemo="---------- 查核項目說明，格式：使用介面、查核項目、查核說明 ---------------------------------" & vbcrlf 
rs.open "select * from [Config] where [Kind]='查核項目' and [Item]<>'' and [Config] not in ('軟體主檔','軟體授權') order by [Mark]",conn
while not rs.eof
  ChkMemo=ChkMemo & rs(2) & "　" & rs(1) & "　" & rs(4) & vbcrlf
  rs.movenext
wend
ChkMemo=ChkMemo & "-------------------------------------------------------------------------------------" & vbcrlf
rs.close

rs1.open "select [成員],[備註] from [View_員工] where [單位]='資訊中心' and [性質]<>'測試' and [備註] like '%" & MailTag & "%' and [備註] not like '%" & LeafTag & "%'",conn	'以後再改為自己異動的不必通知
while not rs1.eof
  '----------mail開頭文字------------------------------------
  LifeAll=rs1(0) & " 長官您好：" & " " & vbcrlf & " " & vbcrlf
  LifeAll=LifeAll & "1. 資源管理系統(IDMS)近一週(" & DT(now-7,"yyyy/mm/dd") & " - " &  DT(now,"yyyy/mm/dd") & ")查核事項自動通知如下所示：" & vbcrlf 
  LifeAll=LifeAll & "2. 介面查詢路徑：http://10.6.1.11:8888  -> [資源管理]  -> [生命履歷]" & vbcrlf & " " & vbcrlf & " " & vbcrlf

  flag=0	'一週通知一次，有異動記錄才需通知
  ChkNum=0	'尚待更正筆數
  UpdNum=0	'異動記錄筆數

  PropNum=0	'財產系統筆數，若<1000則不mail[秘總無此財產編號]
  rs.open "select count(*) from [財產主檔]",conn
  if not rs.eof then PropNum=rs(0)
  rs.close

  '----------查核通知人員(排除僅為 無資產編號 之 相關人員)-------------------------------------
  only_asset_related_flag=0
  
  SQL="SELECT COUNT(*) FROM GetCheckMaintainerAndRelatedToNotify() WHERE [通知人員]='" + rs1(0) + "' OR [通知人員] in (select Kind from Config where Item='" + rs1(0) + "')"
  rs.open SQL,conn  
  'ff.write "#### " & rs1(0) & " = " & rs(0) & " ####" & vbcrlf
  if CInt(rs(0))=0 then	
	'ff.write "SKIP!" & vbcrlf
	only_asset_related_flag=1
  end if
  rs.close
  
  '----------查核通知人員(排除 無資產編號)-------------------------------------
  flag2=0
  
  '----------查核資訊-----------------------------------------
  SQL="select * from [資料查核] where (" + GetSQL("[維護人員]",rs1(0)) + " or " + GetSQL("[相關人員]",rs1(0)) + " or '" + rs1(0) + "' in (select [Item] from [Config] where [Kind]='SSM小組') and [查核結果] like '%非法%')"
  if PropNum<1000 then SQL=SQL & " and [查核結果]<>'秘總無此財產編號'"
  SQL=SQL & " order by [表格名稱],[主鍵編號]"
  rs.open SQL,conn
  if not rs.eof then
    LifeAll=LifeAll & "---------- 資料查核：待更正的資料-------------------------------------------" & vbcrlf
    LifeAll=LifeAll & " 格式：[使用介面]　[主鍵編號]　[資料內容]　[維護人員]　[相關人員]　[查核結果] " & vbcrlf
    LifeAll=LifeAll & "------------------------------------------------------------------------" & " " & vbcrlf
    while not rs.eof
      'if ChkNum<10 then LifeAll=LifeAll & rs(0) & "　" & rs(1) & "　" & rs(2) & "　" & rs(3) & "　" & rs(4) & " " & rs(5) & " "  & vbcrlf
      'ChkNum=ChkNum+1
	  
	  if rs(5)="無資產編號" then		
		'rs2.open "SELECT * FROM Config WHERE Kind='" + rs(3) + "' AND Item='" + rs1(0) + "'",conn		
	    'if ((rs(3)=rs1(0)) or (not rs2.eof)) then 
	    '  if ChkNum<10 then LifeAll=LifeAll & rs(0) & "　" & rs(1) & "　" & rs(2) & "　" & rs(3) & "　" & rs(4) & " " & rs(5) & " "  & vbcrlf			
		'  ChkNum=ChkNum+1
		'end if 
		'rs2.close
	  else
		rs5_no_asset_no = rs(5)
		rs5_no_asset_no = Replace(rs5_no_asset_no, "無資產編號 ", "")
		rs5_no_asset_no = Replace(rs5_no_asset_no, " 無資產編號", "")						
		
		'if ChkNum<10 then LifeAll=LifeAll & rs(0) & "　" & rs(1) & "　" & rs(2) & "　" & rs(3) & "　" & rs(4) & " " & rs(5) & " "  & vbcrlf					
		if ChkNum<10 then LifeAll=LifeAll & rs(0) & "　" & rs(1) & "　" & rs(2) & "　" & rs(3) & "　" & rs(4) & " " & rs5_no_asset_no & " "  & vbcrlf			
		ChkNum=ChkNum+1
		flag2=1
	  end if
		
      rs.movenext
    wend
    LifeAll=LifeAll & " " & vbcrlf
    flag=1
  end if
  rs.close
  if ChkNum>=10 then LifeAll=LifeAll & "********** 查核資訊" & ChkNum & "筆，超過10筆的部分不列出，請自行至資料查核介面查詢完整資料 **********" & vbcrlf & vbcrlf
  if ChkNum>0 then LifeAll=LifeAll & ChkMemo & vbcrlf

  LifeAll=LifeAll & vbcrlf & vbcrlf
  '----------異動記錄------------------------------------------
  LifeAll=LifeAll & "---------- 生命履歷(異動記錄)：僅供參考或複核 --------------------------------------------" & vbcrlf
  SQL=GetSQL("[維護人員]",rs1(0)) + " or " + GetSQL("[原負責人]",rs1(0)) + " or " + GetSQL("[原保管人]",rs1(0)) + " or [主鍵編號] in (select [設備編號] from [View_設備管理] where [設備編號]<>null and [保管人員]='" & rs1(0) & "')"
  call GetLife("實體設備",SQL)
  
  SQL="[主鍵編號] in (select [作業編號] from [View_通用設備] where [作業編號]<> null and (" + GetSQL("[設備維護人員]",rs1(0)) + " or [保管人員]='" + rs1(0) + "'))"
  call GetLife("作業主機",GetSQL("[維護人員]",rs1(0)) + "  or " + SQL)

  call GetLife("*",GetSQL("[維護人員]",rs1(0)))   
  if UpdNum>=10 then LifeAll=LifeAll & "********** 異動記錄" & UpdNum & "筆，超過10筆的部分不列出，請自行至生命履歷介面查詢完整資料 **********" & vbcrlf
  '----------mail相關人員--------------------------------------
  'if flag=1 and (only_asset_related_flag=0 or UpdNum<>0) then
  if flag=1 and ((only_asset_related_flag=0 and flag2=1) or UpdNum<>0) then
      who=rs1(0) : email=rs1(1)
      pos=instr(1,email,MailTag)
      if pos>0 then
        email=mid(email,pos+8,instr(pos+8,email,"#")-pos-8)
        call SendMail(who,email)
		
		ff2.write LifeAll & vbcrlf
		
		'if email<>"" then
		'	ff.write DT(now,"hh:mi:ss") & " " & who & " " & email & " (待更正:" & ChkNum & "筆，異動:" & UpdNum & "筆)" & vbcrlf
		'end if
      end if
  end if
  rs1.movenext
wend
rs1.close

'--------------------------------程式結束------------------------
if Err.Number<>0 then ff.write vbcrlf & Err.Description & " " & vbcrlf
ff.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(End)-------------------------------------------------------------------" & vbcrlf & vbcrlf
ff.close

ff2.write DT(now,"yyyy/mm/dd hh:mi:ss") & "(End)-------------------------------------------------------------------" & vbcrlf & vbcrlf
ff2.close

'msgbox "ok"
'--------------------------------取得生命履歷----------------------------------------------------------------------------------------------------------
Sub GetLife(tbl,SQL)  
  if tbl<>"*" then
    SQL="select * from [生命履歷] where [異動日期]>='" & DT(now-7,"yyyy/mm/dd 00:00:00") & "' and [表格名稱]='" & tbl _
      & "' and (" & SQL & ") order by [表格名稱],[異動人員],[履歷編號]"
  else
    SQL="select * from [生命履歷] where [異動日期]>='" & DT(now-7,"yyyy/mm/dd 00:00:00") & "' and [表格名稱]<>'實體設備' and [表格名稱]<>'作業主機'" _
      & " and (" & SQL & ") order by [表格名稱],[異動人員],[履歷編號]"    
  end if
  'ff.write SQL & vbcrlf

  rs.open SQL,conn
  if not rs.eof then
    while not rs.eof
      LifeLog=rs("生命履歷")	'要先讀，否則讀不出來
      if UpdNum<10 then LifeAll=LifeAll & rs("履歷編號") & "." & rs("異動人員") & "異動 " & tbl & "(" & rs("主鍵編號") & ")" & "　： " & vbcrlf _
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
  ff.write DT(now,"hh:mi:ss") & " " & who & " " & email & " (待更正:" & ChkNum & "筆，異動:" & UpdNum & "筆)" & vbcrlf
 
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
	.Subject  =DT(now," (mm/dd)") & " 資源管理系統(IDMS)資料查核及異動通知(待更正:" & ChkNum & "筆，異動:" & UpdNum & "筆)"
	' .HtmlBody =LifeAll
	.TextBody =LifeAll	'文字模式才無亂碼問題
	.BodyPart.charset = "UTF-8" 

	.Send
  End With 

  Set cdoConfig  =nothing
  Set cdoMessage =nothing
End Sub
'--------------------------------時間格式函數----------------------------------------------------------------------------------------------------------
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
'--------------------------------取得人員SQL函數----------------------------------------------------------------------------------------------------------
Function GetSQL(fld,val)
  GetSQL = fld + "='" + val + "' or " + fld + " in (select Kind from Config where Item='" + val + "')" 
End Function
'--------------------------------更新資料查核資料---------------------------------------------------------------------------------------------------------
Sub CheckDB()
  conn.execute "delete from [資料查核]"
  rs.open "select * from [View_資料查核] where [查核結果]<>''",conn
  while not rs.eof
    conn.execute "insert into [資料查核] values('" & rs(0) & "'," & rs(1) & ",'" & rs(2) & "','" & rs(3) & "','" & rs(4) & "','" & rs(5) & "')"
    rs.movenext
  wend
  rs.close
  ff.write DT(now,"hh:mi:ss") & " 完成更新[資料查核]" & vbcrlf
End Sub
