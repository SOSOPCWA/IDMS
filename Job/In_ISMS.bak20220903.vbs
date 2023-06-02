'============================================================================================================================================================
'功能：每天與ISMS同步更新資產清冊
'============================================================================================================================================================
on error resume next		'寧無作用,不可出錯,以免hung住

'資料庫連結--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
set fs=createobject("scripting.filesystemobject")
set ff=fs.OpenTextFile("In_Isms_" & DT(now,"yyyy") & ".log",8,true)
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
rs.open "SELECT COUNT(*) FROM [View_資產清冊_ISMRASVR3]",conn
Pcount=rs(0)
rs.close

if Pcount>10 and Err.Number=0 then  '資產清冊連不上則不執行
  'ISMS新增資產編號-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Value],[Description] from [View_資產清冊_ISMRASVR3] where [Cancel]=0 and [Number] not in (select [Item] from [Config] where [Kind]='資產清冊')",conn
  while not rs.eof
    Memo=rs(3)	'不這樣取值，會變空白
    conn.execute "insert into [Config] values('資產清冊','" & rs(0) & "','" & rs(1) & "','" & rs(2) & "','" & Memo & "')"
    conn.execute "insert into [生命履歷] values(" & LifeNo & ",'資產清冊',0,'新增：" & rs(0) & " " & rs(1) & " " & Memo & "','SSM小組','','','機房OP','" & DT(now,"yyyy/mm/dd hh:mi:ss") & "','10.6.1.11')"
    LifeNo=LifeNo+1
    rs.movenext
  wend
  rs.close
  
  'ISMS修改資產名稱或描述-----------------------------------------------------------------------------------------------------------------------
  rs.open "select [Number],[Name],[Description],[Config],[Memo],[Value],[Mark] from [View_資產清冊_ISMRASVR3],[Config] where [View_資產清冊_ISMRASVR3].[Number]=[Config].[Item] and [Cancel]=0 and [Kind]='資產清冊' and ([Name]<>[Config] or [Description]<>[Memo])",conn
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
  rs.open "select * from [Config] where [Kind]='資產清冊' and [Config] not like '(已註銷)%' and [Item] not in (select [Number] from [View_資產清冊_ISMRASVR3] where [Cancel]=0)",conn,3,3
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