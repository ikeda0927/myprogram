Dim ans,wmp
ans=MsgBox("CD�܂���DVD�����܂����H",vbYesNo)
If ans=vbYes Then
 Set wmp = CreateObject("WMPlayer.OCX")
 wmp.cdromcollection.item(0).eject()
 WScript.Quit(eLevel)
Else

End If
