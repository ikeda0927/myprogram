Dim ans,wmp
ans=MsgBox("CDまたはDVDを入れますか？",vbYesNo)
If ans=vbYes Then
 Set wmp = CreateObject("WMPlayer.OCX")
 wmp.cdromcollection.item(0).eject()
 WScript.Quit(eLevel)
Else

End If
