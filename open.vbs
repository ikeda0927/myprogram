Dim ans,wmp
ans=MsgBox("CD‚Ü‚½‚ÍDVD‚ğ“ü‚ê‚Ü‚·‚©H",vbYesNo)
If ans=vbYes Then
 Set wmp = CreateObject("WMPlayer.OCX")
 wmp.cdromcollection.item(0).eject()
 WScript.Quit(eLevel)
Else

End If
