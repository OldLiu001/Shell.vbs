<<<<<<< HEAD:Libraries/VBScript/Music.vbs
'Play(MusicPath)

Set WMP=createObject("WMPlayer.OCX") 

Sub Music_Open(MusicPath)
		wmp.URL=MusicPath
		Do Until wmp.playState<>1
			WScript.Sleep(500)
		Loop
End Sub

Sub Music_Pause()
	wmp.controls.pause
End Sub

Sub Music_Play()
	wmp.controls.play
End Sub

Sub Music_Stop()
	wmp.controls.stop
End Sub
=======
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
>>>>>>> parent of e1920d8... Merge pull request #5 from OldLiu001/OldLiu:Lib/VBSLib/Music.vbs
