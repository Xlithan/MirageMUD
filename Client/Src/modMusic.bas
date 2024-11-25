Attribute VB_Name = "Music"
Option Explicit
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public bInit_Music As Boolean
Public bInit_Sound As Boolean
Public curSong As String

Private songHandle As Long
Private streamHandle As Long

Public CurrentMusic As String

Public Function Init_Music() As Boolean
Dim result As Boolean

    On Error GoTo errorhandler
    
    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)
    If Not result Then GoTo errorhandler
    
    ' return positive
    Init_Music = True
    bInit_Music = True
    
    ' init FMOD sound system
    FSOUND_Init 44100, 32, 0
    bInit_Sound = True
   
    Exit Function
    
errorhandler:
    Init_Music = False
    bInit_Music = False
    bInit_Sound = False
End Function

Public Sub Destroy_Music()
    ' destroy music engine
    Stop_Music
    FSOUND_Close
    bInit_Music = False
    curSong = vbNullString
End Sub

Public Sub Play_Music(ByVal song As String)
    If Not bInit_Music Then Exit Sub
    
    ' does it exist?
    If Not FileExist(MUSIC_PATH & song) Then Exit Sub
    
    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    
    ' stop the existing music
    Stop_Music
    
    ' find the extension
    Select Case Right$(song, 4)
        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.Path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, 150
            
        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.Path & MUSIC_PATH & song, FSOUND_LOOP_NORMAL, 0, 0)
            ' play it
            FSOUND_Stream_Play 0, streamHandle
            ' set volume
            FSOUND_SetVolume streamHandle, 150
        Case Else
            Exit Sub
    End Select
    
    ' new current song
    curSong = song
End Sub

Public Sub Stop_Music()
    If Not streamHandle = 0 Then
        ' stop stream
        FSOUND_Stream_Stop streamHandle
        ' destroy
        FSOUND_Stream_Close streamHandle
        streamHandle = 0
    End If
    
    If Not songHandle = 0 Then
        ' stop song
        FMUSIC_StopSong songHandle
        ' destroy
        FMUSIC_FreeSong songHandle
        songHandle = 0
    End If
    
    ' no music
    curSong = vbNullString
End Sub

Public Sub Play_Sound(Sound As String)
Dim Handle As Long

    If GameData.Sound = 0 Then Exit Sub
    If FileExist(SFX_PATH & Sound) = False Then Exit Sub
    
    Call sndPlaySound(App.Path & SFX_PATH & Sound, SND_ASYNC Or SND_NODEFAULT)
End Sub

Public Sub StopSound()
    Dim X As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound("", wFlags)
End Sub
