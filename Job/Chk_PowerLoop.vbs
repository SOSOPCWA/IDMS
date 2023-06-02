'-----------------------------------------------------------------
InDB="sos-vm16" 

strConn="Driver={SQL SERVER};Server=" & InDB & ";Trusted_Connection=True;Database=" 
Dim connIDMS    : Set connIDMS=CreateObject("ADODB.Connection")   : connIDMS.Open strConn & "IDMS"

'------------------取得起始編號-------------------------------------
'on error resume next
dim i : i=0
dim str : str=""

call TreePower(0)

Sub TreePower(DevNo)
  dim rs : set rs=createobject("ADODB.Recordset")
  rs.open "select [下游編號] from [接電] where [上游編號]=" & DevNo,connIDMS
  while not rs.eof
    if instr(1,str,"(" & DevNo & "," & rs(0) & ")")>0 then
      msgbox "(" & DevNo & "," & rs(0) & ")重複，將產生無窮迴圈"
      WScript.Quit
    end if
    str=str & ",(" & DevNo & "," & rs(0) & ")"
    call TreePower(rs(0))
    rs.movenext
    i=i+1
  wend
  rs.close
End Sub

msgbox i & "筆OK! 沒有無窮迴圈"

