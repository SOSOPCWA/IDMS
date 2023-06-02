'------- 將機房所有坐標定位之，定位編號以負值表示
'------- Y<0 代表非機房之坐標 ------------------------------------------------------
'-----------------------------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=sos-vm16;Trusted_Connection=True;Database=" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn & "IDMS"

conn.execute "delete from [定位設定] where [定位方式]='坐標'"
PointerNo=-1

for x=-1 to 65	'機房外圍
  for y=1 to 31
    conn.execute "insert into [定位設定] values(" & PointerNo & ",'坐標','機房外圍','(" & x & "," & y & ")',1," & x & "," & y & "," & x & "," & y & "," & x & "," & y & ")"
    PointerNo=PointerNo-1
  next
next

for x=1 to 10	'前機房
  for y=2 to 29
    conn.execute "update [定位設定] set [區域名稱]='前機房' where [定位方式]='坐標' and [坐標X]=" & x & " and [坐標Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=11 to 37	'第一機房A
  for y=2 to 17
    conn.execute "update [定位設定] set [區域名稱]='第一機房' where [定位方式]='坐標' and [坐標X]=" & x & " and [坐標Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=11 to 29	'第一機房B
  for y=18 to 29
    conn.execute "update [定位設定] set [區域名稱]='第一機房' where [定位方式]='坐標' and [坐標X]=" & x & " and [坐標Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=41 to 64	'第二機房
  for y=12 to 29
    conn.execute "update [定位設定] set [區域名稱]='第二機房' where [定位方式]='坐標' and [坐標X]=" & x & " and [坐標Y]=" & y
    PointerNo=PointerNo-1
  next
next


'部分第一機房改為機房外圍
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=31 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=32 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=32 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=33 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=33 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=33 and [坐標Y]=4"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=34 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=34 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=34 and [坐標Y]=4"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=34 and [坐標Y]=5"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=35 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=35 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=35 and [坐標Y]=4"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=35 and [坐標Y]=5"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=35 and [坐標Y]=6"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=4"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=5"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=6"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=36 and [坐標Y]=7"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=2"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=3"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=4"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=5"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=6"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=7"
conn.execute "update [定位設定] set [區域名稱]='機房外圍' where [定位方式]='坐標' and [坐標X]=37 and [坐標Y]=8"

msgbox "ok!"