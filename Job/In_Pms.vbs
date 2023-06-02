'============================================================================================================================================================
'功能：每天匯入財產系統資料至IDMS，以加速查詢及帶入功能，並與ISMS同步更新資產清冊
'============================================================================================================================================================
on error resume next		'寧無作用,不可出錯,以免hung住

'資料庫連結--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
set fs=createobject("scripting.filesystemobject")
set ff=fs.OpenTextFile("In_Pms_" & DT(now,"yyyy") & ".log",8,true)
NowDT=DT(now,"yyyy/mm/dd hh:mi:ss")
ff.write NowDT & " "

strConn="Driver={SQL SERVER};Server=10.6.1.11;Trusted_Connection=True;Database=" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn & "IDMS" : conn.CommandTimeout = 1800
set rs=createobject("ADODB.Recordset")

LifeNo=1
rs.open "select max([履歷編號]) from [生命履歷]",conn
if not rs.eof then LifeNo=rs(0)+1
rs.close  

'與ISMS同步更新資產清冊-----------------------------------------------------------------------------------------------------------------------
Pcount=0
rs.open "SELECT COUNT(*) FROM [View_資產清冊]",conn
Pcount=rs(0)
rs.close

if Pcount>10 and Err.Number=0 then  '資產清冊連不上則不執行
  'ISMS新增資產編號-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Value],[Description] from [View_資產清冊] where [Cancel]=0 and [Number] not in (select [Item] from [Config] where [Kind]='資產清冊')",conn
  while not rs.eof
    Memo=rs(3)	'不這樣取值，會變空白
    conn.execute "insert into [Config] values('資產清冊','" & rs(0) & "','" & rs(1) & "','" & rs(2) & "','" & Memo & "')"
    conn.execute "insert into [生命履歷] values(" & LifeNo & ",'資產清冊',0,'新增：" & rs(0) & " " & rs(1) & " " & Memo & "','SSM小組','','','機房OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close
  'ISMS修改資產名稱或描述-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Description],[Config],[Memo],[Value],[Mark] from [View_資產清冊],[Config] where [View_資產清冊].[Number]=[Config].[Item] and [Cancel]=0 and [Kind]='資產清冊' and ([Name]<>[Config] or [Description]<>[Memo])",conn
  while not rs.eof  '資產價值的異動判斷包在Memo中，不另外處理
    Memo=rs(2)	'不這樣取值，會變空白
    Life="修改：" & rs(0) & " " & rs(3) & "->" & rs(1) & " " & rs(4) & "->" & Memo
    conn.execute "Update [Config] set [Config]='" & rs(1) & "',[Mark]='" & rs(5) & "',[Memo]='" & Memo & "' where [Kind]='資產清冊' and [Item]='" & rs(0) & "'"
    conn.execute "insert into [生命履歷] values(" & LifeNo & ",'資產清冊',0,'修改：" & rs(0) & " " & rs(3) & "->" & rs(1) & " " & rs(4) & "->" & Memo & "','SSM小組','','','機房OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close

  'ISMS刪除資產編號-----------------------------------------------------------------------------------------------------------------------
  rs.open "select * from [Config] where [Kind]='資產清冊' and [Config] not like '(已註銷)%' and [Item] not in (select [Number] from [View_資產清冊] where [Cancel]=0)",conn,3,3
  while not rs.eof    
    conn.execute "insert into [生命履歷] values(" & LifeNo & ",'資產清冊',0,'刪除：" & rs(0) & " " & rs(1) & " " & rs(2) & " (請於IDMS系統設定手動刪除，並注意資料相依性)','SSM小組','','','機房OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    rs(2)="(已註銷)" & rs(2)
    rs.update
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close

  ff.write "資產清冊：" & DT(now,"hh:mi:ss") & " "
end if

'匯入秘總財產-----------------------------------------------------------------------------------------------------------------------
conn.execute "drop table [秘總財產]"
conn.execute "select * into [秘總財產] from [View_秘總財產]"
ff.write "秘總財產：" & DT(now,"hh:mi:ss") & " "    

'秘總財產異動查核-----------------------------------------------------------------------------------------------------------------------
if Err.Number=0 then    '秘總財產連結不成功則不執行
  Pcount=0
  rs.open "SELECT COUNT(*) FROM [秘總財產]",conn
  Pcount=rs(0)
  rs.close

  if Pcount>5000 then  '秘總財產匯入不成功則不執行
    '檢查秘總財產保管人變更-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [設備編號],[財產主檔].[財產編號A] as [財產編號A],[財產主檔].[財產編號B] as [財產編號B],[財產主檔].[財產別名] as [財產別名]" _
      & ",[維護人員],[秘總財產].[保管人員] as [保管人員],[財產主檔].[保管人員] as [原保管人]" _
      & " from [財產主檔],[秘總財產],[實體設備]" _
      & " where [財產主檔].[財產編號A]=[秘總財產].[財產編號A] and [財產主檔].[財產編號B]=[秘總財產].[財產編號B]" _
      & " and   [財產主檔].[財產編號A]=[實體設備].[財產編號A] and [財產主檔].[財產編號B]=[實體設備].[財產編號B]" _
      & " and [財產主檔].[保管人員]<>[秘總財產].[保管人員]",conn
    while not rs.eof 
      conn.execute "insert into [生命履歷] values(" & LifeNo & ",'秘總財產'," & rs(0) _
        & ",'秘總財產保管人變更(" & rs(1) & " " & rs(2) & " " & rs(3) & ")：" & rs(6) & "->" & rs(5) & "'" _
        & ",'" & rs(4) & "','財產通知','" & rs(6) & "','財管人員','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '檢查秘總財產報廢-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [財產主檔].[財產編號A] + ' ' + [財產主檔].[財產編號B] as [財產編號],[財產主檔].[財產別名]" _
      & ",[財產主檔].[廠牌],[財產主檔].[型式],[財產主檔].[購買日期],[財產主檔].[保管人員]" _
      & " from [財產主檔],[秘總財產] where [財產主檔].[財產編號A]=[秘總財產].[財產編號A] and [財產主檔].[財產編號B]=[秘總財產].[財產編號B]" _
      & " and [財產主檔].[無效註記]='' and [秘總財產].[無效註記]='N'",conn    
    while not rs.eof 
      conn.execute "insert into [生命履歷] values(" & LifeNo & ",'秘總財產',0" _
        & ",'秘總財產報廢" & "(" & rs(0) & " " & rs(1) & ")：" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','財產通知','','財管人員','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '檢查秘總財產新增-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [財產編號A] + ' ' + [財產編號B] as [財產編號],[財產別名],[廠牌],[型式],[購買日期],[保管人員] from [秘總財產]" _
      & " where [財產編號A] + ' ' + [財產編號B] not in (select [財產編號A] + ' ' + [財產編號B] from [財產主檔])",conn
    while not rs.eof     
      conn.execute "insert into [生命履歷] values(" & LifeNo & ",'秘總財產',0" _
        & ",'秘總財產新增(" & rs(0) & " " & rs(1) & ")：" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','財產通知','','財管人員','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    '檢查秘總財產移除-----------------------------------------------------------------------------------------------------------------------
    rs.open "select [財產編號A] + ' ' + [財產編號B] as [財產編號],[財產別名],[廠牌],[型式],[購買日期],[保管人員] from [財產主檔]" _
      & " where [財產編號A] + ' ' + [財產編號B] not in (select [財產編號A] + ' ' + [財產編號B] from [秘總財產])",conn
    while not rs.eof 
      conn.execute "insert into [生命履歷] values(" & LifeNo & ",'秘總財產',0" _
        & ",'秘總財產移除(" & rs(0) & " " & rs(1) & ")：" & rs(2) & " " & rs(3) & " " & rs(4) & " " & rs(5) & "'" _
        & ",'" & rs(5) & "','財產通知','','財管人員','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
      LifeNo=LifeNo+1
      rs.movenext
    wend
    rs.close 
    ff.write "異動查核：" & DT(now,"hh:mi:ss") & " " 

    '更新財產主檔-----------------------------------------------------------------------------------------------------------------------
    conn.execute "drop table [財產主檔]"
    conn.execute "select * into [財產主檔] from [秘總財產]" 
    conn.execute "ALTER TABLE [財產主檔] ADD  CONSTRAINT [PK_財產主檔] PRIMARY KEY CLUSTERED ([機關代號] ASC,[財產編號A] ASC,[財產編號B] ASC)" _
        & " WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,SORT_IN_TEMPDB=OFF,IGNORE_DUP_KEY=OFF,ONLINE=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    conn.execute "UPDATE [財產主檔] SET [財產別名]='' where [財產別名] is null"
    conn.execute "UPDATE [財產主檔] SET [廠牌]='' where [廠牌] is null"
    conn.execute "UPDATE [財產主檔] SET [型式]='' where [型式] is null"
    conn.execute "UPDATE [財產主檔] SET [計量單位]='' where [計量單位] is null"
    conn.execute "UPDATE [財產主檔] SET [使用年限]=0 where [使用年限] is null"
    conn.execute "UPDATE [財產主檔] SET [員工代號]='' where [員工代號] is null"
    conn.execute "UPDATE [財產主檔] SET [財產地點]='' where [財產地點] is null"
    conn.execute "UPDATE [財產主檔] SET [保管課別]='' where [保管課別] is null"
    conn.execute "UPDATE [財產主檔] SET [保管人員]='' where [保管人員] is null"
    conn.execute "UPDATE [財產主檔] SET [無效註記]='' where [無效註記] is null"
    conn.execute "UPDATE [財產主檔] SET [財產附屬]='' where [財產附屬] is null"

    ff.write "財產主檔：" & DT(now,"hh:mi:ss") & " " 

  end if  '秘總財產匯入不成功則不執行
end if    '秘總財產連結不成功則不執行(Err.Number)
'------------------------------------------------------------
if Err.Number<>0 then ff.write Err.Description & " "

ff.write DT(now,"hh:mi:ss") & " (" & Pcount & ")" & vbcrlf
ff.close

'msgbox "ok!"
'--------------------------------時間格式函數---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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