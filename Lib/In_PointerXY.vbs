'------- �N���ЩҦ����Щw�줧�A�w��s���H�t�Ȫ��
'------- Y<0 �N��D���Ф����� ------------------------------------------------------
'-----------------------------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=sos-vm16;Trusted_Connection=True;Database=" 
Dim conn    : Set conn=CreateObject("ADODB.Connection")   : conn.Open strConn & "IDMS"

conn.execute "delete from [�w��]�w] where [�w��覡]='����'"
PointerNo=-1

for x=-1 to 65	'���Х~��
  for y=1 to 31
    conn.execute "insert into [�w��]�w] values(" & PointerNo & ",'����','���Х~��','(" & x & "," & y & ")',1," & x & "," & y & "," & x & "," & y & "," & x & "," & y & ")"
    PointerNo=PointerNo-1
  next
next

for x=1 to 10	'�e����
  for y=2 to 29
    conn.execute "update [�w��]�w] set [�ϰ�W��]='�e����' where [�w��覡]='����' and [����X]=" & x & " and [����Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=11 to 37	'�Ĥ@����A
  for y=2 to 17
    conn.execute "update [�w��]�w] set [�ϰ�W��]='�Ĥ@����' where [�w��覡]='����' and [����X]=" & x & " and [����Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=11 to 29	'�Ĥ@����B
  for y=18 to 29
    conn.execute "update [�w��]�w] set [�ϰ�W��]='�Ĥ@����' where [�w��覡]='����' and [����X]=" & x & " and [����Y]=" & y
    PointerNo=PointerNo-1
  next
next

for x=41 to 64	'�ĤG����
  for y=12 to 29
    conn.execute "update [�w��]�w] set [�ϰ�W��]='�ĤG����' where [�w��覡]='����' and [����X]=" & x & " and [����Y]=" & y
    PointerNo=PointerNo-1
  next
next


'�����Ĥ@���Чאּ���Х~��
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=31 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=32 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=32 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=33 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=33 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=33 and [����Y]=4"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=34 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=34 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=34 and [����Y]=4"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=34 and [����Y]=5"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=35 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=35 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=35 and [����Y]=4"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=35 and [����Y]=5"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=35 and [����Y]=6"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=4"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=5"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=6"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=36 and [����Y]=7"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=2"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=3"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=4"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=5"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=6"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=7"
conn.execute "update [�w��]�w] set [�ϰ�W��]='���Х~��' where [�w��覡]='����' and [����X]=37 and [����Y]=8"

msgbox "ok!"