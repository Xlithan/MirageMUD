Attribute VB_Name = "Database"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Function FileExist(ByVal filename As String, _
                          Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub AddLog(ByVal Text As String)

    Dim filename As String

    Dim F        As Long
    
    If DebugMode Then
        If Not frmDebug.Visible Then
            frmDebug.Visible = True
        End If
        
        filename = App.Path & LOG_PATH & LOG_DEBUG
        
        If Not FileExist(LOG_DEBUG, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If
        
        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

Public Sub SaveRoom(ByVal RoomNum As Long)

    Dim filename As String

    Dim F        As Long
    
    filename = App.Path & ROOM_PATH & "room" & RoomNum & ROOM_EXT
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Room
    Close #F
End Sub

Public Sub LoadRooms(ByVal i As Long)

    Dim BlankRoom As RoomRec
    
    If FileExist(ROOM_PATH & "room" & i & ROOM_EXT) Then
        Call LoadRoom(i)
        
        If GameData.Music = 1 Then
            If Len(Room.Music) > 0 Then
                If Trim$(CStr(Room.Music)) <> CurrentMusic Then
                    Stop_Music
                    Play_Music (Trim$(CStr(Room.Music)))
                    CurrentMusic = Trim$(CStr(Room.Music))
                End If
    
            Else
                Stop_Music
                CurrentMusic = 0
            End If
        End If

    End If

End Sub

Public Sub LoadRoom(ByVal RoomNum As Long)

    Dim filename As String

    Dim F        As Long
    
    filename = App.Path & ROOM_PATH & "room" & RoomNum & ROOM_EXT
    
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Room
    Close #F
End Sub

Public Sub LoadDataFile()

    Dim filename As String

    Dim F        As Long
    
    ' Check if the logs directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\data\logs")
    End If
    
    ' Check if the music directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data\music", vbDirectory)) <> "music" Then
        Call MkDir(App.Path & "\data\music")
    End If
    
    filename = App.Path & DATA_PATH & "config.dat"
    
    If Not FileExist("data\config.dat") Then
        GameData.IP = "127.0.0.1"
        GameData.Port = 7777
        F = FreeFile
        Open filename For Binary As #F
        Put #F, , GameData
        Close #F
    Else
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , GameData
        Close #F
    End If

End Sub

Public Sub CheckAvatars()

    Dim i As Long
    
    i = 1
    
    While FileExist(GFX_PATH & "avatars\Players\" & i & GFX_EXT)

        NumAvatars = NumAvatars + 1
        i = i + 1

    Wend
    
End Sub

Public Sub CheckNPCAvatars()

    Dim i As Long
    
    i = 1
    
    While FileExist(GFX_PATH & "avatars\NPCs\" & i & GFX_EXT)

        NumNPCAvatars = NumNPCAvatars + 1
        i = i + 1

    Wend
    
End Sub

Public Sub CheckSpells()

    Dim i As Long
    
    i = 1
    
    While FileExist(GFX_PATH & "Spells\" & i & GFX_EXT)

        NumSpells = NumSpells + 1
        i = i + 1

    Wend
    
    ReDim Tr_Spell(1 To NumSpells)
    
    ReDim SpellTimer(1 To NumSpells)
    
End Sub

Public Sub CheckItems()

    Dim i As Long
    
    i = 1
    
    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)

        numitems = numitems + 1
        i = i + 1

    Wend
    
    ReDim DDS_Item(1 To numitems)
    ReDim DDSD_Item(1 To numitems)
    
    ReDim ItemTimer(1 To numitems)
    
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).name = vbNullString
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
End Sub

Public Sub ClearItems()

    Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Public Sub ClearRoomItem(ByVal Index As Long, ByVal RoomNum As Long)
    Call ZeroMemory(ByVal VarPtr(RoomItem(Index, RoomNum)), LenB(RoomItem(Index, RoomNum)))
End Sub

Public Sub ClearRoom()
    Call ZeroMemory(ByVal VarPtr(Room), LenB(Room))
    Room.name = vbNullString
End Sub

Public Sub ClearRoomItems()

    Dim i As Long

    Dim n As Long
    
    For i = 1 To MAX_ROOM_ITEMS
        For n = 1 To MAX_ROOMS
            Call ClearRoomItem(i, n)
        Next n
    Next
    
    Call RefreshEntityList(MyIndex, 2)
End Sub

Public Sub ClearRoomNpc(ByVal Index As Long, ByVal Room As Long)
    Call ZeroMemory(ByVal VarPtr(RoomNpc(Index, Room)), LenB(RoomNpc(Index, Room)))
End Sub

Public Sub ClearRoomNpcs()

    Dim i As Long

    Dim n As Long
    
    For i = 1 To MAX_ROOM_NPCS
        For n = 1 To MAX_ROOMS
            Call ClearRoomNpc(i, n)
        Next
    Next
    
    Call RefreshEntityList(MyIndex, 1)
End Sub

' *****************************
' ** Player Public Functions **
' *****************************

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).name)
End Function

Public Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Public Function GetPlayerAvatar(ByVal Index As Long) As Long
    GetPlayerAvatar = Player(Index).Avatar

    If GetPlayerAvatar = 0 Then GetPlayerAvatar = 1
End Function

Public Sub SetPlayerAvatar(ByVal Index As Long, ByVal Avatar As Long)
    Player(Index).Avatar = Avatar
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Public Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Public Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.Strength) + GetPlayerStat(Index, Stats.Defense) + GetPlayerStat(Index, Stats.Magic) + GetPlayerStat(Index, Stats.speed) + GetPlayerPOINTS(Index)) * 25
End Function

Public Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Public Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Public Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

End Sub

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    Select Case Vital

        Case HP
            GetPlayerMaxVital = Player(Index).MaxHP

        Case MP
            GetPlayerMaxVital = Player(Index).MaxMP

        Case SP
            GetPlayerMaxVital = Player(Index).MaxSP

        Case Stamina
            GetPlayerMaxVital = Player(Index).MaxStamina
    End Select

End Function

Public Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Public Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Public Function GetPlayerRoom(ByVal Index As Long) As Long
    GetPlayerRoom = Player(Index).Room
End Function

Public Sub SetPlayerRoom(ByVal Index As Long, ByVal RoomNum As Long)
    Player(Index).Room = RoomNum
End Sub

Public Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Public Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = PlayerInv(InvSlot).num
End Function

Public Sub SetPlayerInvItemNum(ByVal Index As Long, _
                               ByVal InvSlot As Long, _
                               ByVal ItemNum As Long)
    PlayerInv(InvSlot).num = ItemNum
End Sub

Public Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = PlayerInv(InvSlot).Value
End Function

Public Sub SetPlayerInvItemValue(ByVal Index As Long, _
                                 ByVal InvSlot As Long, _
                                 ByVal ItemValue As Long)
    PlayerInv(InvSlot).Value = ItemValue
End Sub

Public Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = PlayerInv(InvSlot).Dur
End Function

Public Sub SetPlayerInvItemDur(ByVal Index As Long, _
                               ByVal InvSlot As Long, _
                               ByVal ItemDur As Long)
    PlayerInv(InvSlot).Dur = ItemDur
End Sub

Public Function GetPlayerEquipmentSlot(ByVal Index As Long, _
                                       ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Equipment(EquipmentSlot)
End Function

Public Sub SetPlayerEquipmentSlot(ByVal Index As Long, _
                                  ByVal InvNum As Long, _
                                  ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

Public Function CheckRoomRevision(ByVal RoomNum As Long, ByVal rev As Long) As Boolean
    Call LoadRoom(RoomNum)

    If Room.Revision = rev Then
        CheckRoomRevision = True
    End If

End Function
