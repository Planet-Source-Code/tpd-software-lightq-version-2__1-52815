Attribute VB_Name = "Game_Sound"
' ######################################
' ##
' ##  LightQ DirectX multichannel
' ##  sound system
' ##

Public AudioPresent As Boolean

Public SoundEnabled As Boolean
Public MusicEnabled As Boolean

Public MusicPlaying As Boolean

Private Const NUM_CHANNELS = 8   ' 8 channels
Private Const Volume = 0         ' max volume

Private dx As New DirectX8
Private ds As DirectSound8

Private dmParams As DMUS_AUDIOPARAMS
Private hEvent   As Long

' loads music, it transfers
' the contents of a file into memory
Private mLoader As DirectMusicLoader8

' controls the music
Private mPerformance As DirectMusicPerformance8

' stores the music in memory
Private mSegment As DirectMusicSegment8

' create multiple sound buffers
Private dsBuffer(NUM_CHANNELS) As DirectSoundSecondaryBuffer8

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Sub InitDirectAudio()
    
    On Error GoTo InitError
   
    Dim sTag As String
   
    AudioPresent = False
    
    With frmScreen
    
       If waveOutGetNumDevs() > 0 Then
    
          Set ds = dx.DirectSoundCreate(vbNullString)
    
          ds.SetCooperativeLevel .hwnd, DSSCL_PRIORITY

          AudioPresent = True
       
       End If

       .chk_settings(0).Enabled = AudioPresent
       .sVol.Enabled = AudioPresent
       .chk_settings(1).Enabled = AudioPresent
       .mVol.Enabled = AudioPresent
       
    End With
    
    Exit Sub

InitError:
ShowError "Could not initialise direct audio"
    
End Sub

Sub ClearDirectAudio()

    ' clear all open channels and instances
    Dim iX As Long
    
    For iX = 0 To UBound(dsBuffer)
       If Not dsBuffer(iX) Is Nothing Then dsBuffer(iX).Stop
       Set dsBuffer(iX) = Nothing
    Next iX
    
    If Not mPerformance Is Nothing Then
       mPerformance.CloseDown
       Set mPerformance = Nothing
    End If
    
    If ObjPtr(mSegment) Then
       Set mSegment = Nothing
    End If
    
    If ObjPtr(mLoader) Then
       Set mLoader = Nothing
    End If
    
    If Not ds Is Nothing Then
       Set ds = Nothing
    End If
    
    If ObjPtr(dx) Then
       Set dx = Nothing
    End If
    
End Sub

Sub PlaySound(Filename As String, Optional cVol As Long = Volume)
    
    On Error GoTo PlayError
    
    ' don't play when no soundcard present
    If Not AudioPresent Then Exit Sub

    ' only play when sound is on
    If Not SoundEnabled Then Exit Sub
    
    Dim bufferDesc As DSBUFFERDESC
    Dim iX         As Long
    
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    
    ' cleanup completed channels and find a free channel
    For iX = 0 To UBound(dsBuffer)
       
       If Not dsBuffer(iX) Is Nothing Then
         
         If dsBuffer(iX).GetStatus = 0 Or dsBuffer(iX).GetStatus = DSBSTATUS_TERMINATED Then
            Set dsBuffer(iX) = Nothing
         End If
         
       End If
       
       If dsBuffer(iX) Is Nothing Then
          
          Set dsBuffer(iX) = ds.CreateSoundBufferFromFile(App.Path & "\data\sound\" & Filename & ".wav", bufferDesc)
          
          dsBuffer(iX).SetCurrentPosition 0
          dsBuffer(iX).SetVolume frmScreen.sVol                   ' Range: -10000 TO 0
          dsBuffer(iX).SetPan DSBPAN_CENTER                       ' Range: -10000 TO 10000
          dsBuffer(iX).Play DSBPLAY_DEFAULT
          Exit For
          
      End If
       
    Next iX
    
    Exit Sub
    
PlayError:
ShowError "Could not play " & LCase(Filename) & ".wav"

End Sub

Sub FreeSoundBuffers()

    ' don't do when no soundcard present
    If Not AudioPresent Then Exit Sub

    ' only do when sound is on
    If Not SoundEnabled Then Exit Sub

    Dim iX         As Long
    
    For iX = 0 To UBound(dsBuffer)
       
       If Not dsBuffer(iX) Is Nothing Then Set dsBuffer(iX) = Nothing
    
    Next iX

End Sub

Sub PlayMusic(Filename As String)

    On Error GoTo PlayError
    
    ' don't do when no soundcard present
    If Not AudioPresent Then Exit Sub

    ' don't play when no soundcard present
    If Not MusicEnabled Then Exit Sub
    
    Set mLoader = dx.DirectMusicLoaderCreate
    ' create the performance
    Set mPerformance = dx.DirectMusicPerformanceCreate
    ' start up the performance, telling DirectMusic
    ' the handle of the form
    mPerformance.InitAudio frmScreen.hwnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    ' tell DirectMusic to do all the sound downloading
    ' stuff itself because we can't be bothered to do it
    mPerformance.SetMasterAutoDownload True
    
    mLoader.SetSearchDirectory App.Path & "\data\music"
        
    Set mSegment = mLoader.LoadSegment(Filename & ".mid")
    
    mSegment.SetStandardMidiFile
 
    ' infinite loop
    mSegment.SetRepeats -1

    SetMusicVolume

    mPerformance.PlaySegmentEx mSegment, DMUS_SEGF_DEFAULT, 0
    
    MusicPlaying = True
    
    Exit Sub
    
PlayError:
ShowError "Could not play " & LCase(Filename) & ".mid"
 
End Sub

Sub SetMusicVolume(Optional GameOverride As Boolean = False, Optional Volume As Long)

    ' don't do when no soundcard present
    If Not AudioPresent Then Exit Sub

    ' nothing to set
    If Not MusicEnabled Or mPerformance Is Nothing Then Exit Sub
    
    If GameOverride Then
       mPerformance.SetMasterVolume Volume               ' -10000 to 10000  (0 normal)
    Else
       mPerformance.SetMasterVolume frmScreen.mVol       ' -10000 to 10000  (0 normal)
    End If

End Sub

Sub StopMusic()

    ' don't do when no soundcard present
    If Not AudioPresent Then Exit Sub
   
    ' nothing to stop
    If mPerformance Is Nothing Then Exit Sub

    mPerformance.StopEx mSegment, 0, DMUS_SEGF_DEFAULT
 
    MusicPlaying = False
 
End Sub
