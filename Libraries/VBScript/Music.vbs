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