'Play(MusicPath)

Set WMP=createObject("WMPlayer.OCX") 

Sub Music_Play(MusicPath)
		wmp.URL=MusicPath
		MsgBox("Playing "&MusicPath)
		Do Until wmp.playState<>1
			WScript.Sleep(500)
		Loop
End Sub

Sub Music_Pause()
	wmp.controls.pause
End Sub

Sub Music_Continue()
	wmp.controls.play
End Sub

Function Music_Time()
	Music_Time=wmp.currentMedia.durationString()
End Function

Function Music_Info()
	wmp.currentMedia.getItemInfo(Music_Info)
End Function

Function Music_ProcessRate()
	Music_ProcessRate=wmp.controls.currentPositionString&"/"&wmp.currentMedia.durationString
End Function
