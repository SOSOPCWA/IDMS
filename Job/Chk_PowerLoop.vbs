'-----------------------------------------------------------------
InDB="sos-vm16" 

strConn="Driver={SQL SERVER};Server=" & InDB & ";Trusted_Connection=True;Database=" 
Dim connIDMS    : Set connIDMS=CreateObject("ADODB.Connection")   : connIDMS.Open strConn & "IDMS"

'------------------���o�_�l�s��-------------------------------------
'on error resume next
dim i : i=0
dim str : str=""

call TreePower(0)

Sub TreePower(DevNo)
  dim rs : set rs=createobject("ADODB.Recordset")
  rs.open "select [�U��s��] from [���q] where [�W��s��]=" & DevNo,connIDMS
  while not rs.eof
    if instr(1,str,"(" & DevNo & "," & rs(0) & ")")>0 then
      msgbox "(" & DevNo & "," & rs(0) & ")���ơA�N���͵L�a�j��"
      WScript.Quit
    end if
    str=str & ",(" & DevNo & "," & rs(0) & ")"
    call TreePower(rs(0))
    rs.movenext
    i=i+1
  wend
  rs.close
End Sub

msgbox i & "��OK! �S���L�a�j��"

