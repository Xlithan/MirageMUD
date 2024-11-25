Attribute VB_Name = "DataHandling"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SAllChars) = GetAddress(AddressOf HandleAllChars)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerSp) = GetAddress(AddressOf HandlePlayerSp)
    HandleDataSub(SPlayerStamina) = GetAddress(AddressOf HandlePlayerStamina)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerExp) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForRoom) = GetAddress(AddressOf HandleCheckForRoom)
    HandleDataSub(SRoomData) = GetAddress(AddressOf HandleRoomData)
    HandleDataSub(SRoomItemData) = GetAddress(AddressOf HandleRoomItemData)
    HandleDataSub(SRoomNpcData) = GetAddress(AddressOf HandleRoomNpcData)
    HandleDataSub(SRoomDone) = GetAddress(AddressOf HandleRoomDone)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SRoomMsg) = GetAddress(AddressOf HandleRoomMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(SEditRoom) = GetAddress(AddressOf HandleEditRoom)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(SREditor) = GetAddress(AddressOf HandleRefresh)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SCastSpell) = GetAddress(AddressOf HandleSpellCast)
    HandleDataSub(SSendMaxes) = GetAddress(AddressOf HandleMaxes)
    HandleDataSub(SSync) = GetAddress(AddressOf HandleSync)
    HandleDataSub(SRoomRevs) = GetAddress(AddressOf HandleRoomRevs)
End Sub

Sub HandleData(ByRef Data() As Byte)

    Dim Buffer  As clsBuffer

    Dim MsgType As Integer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadInteger
    
    If MsgType < 0 Or MsgType >= SMSG_COUNT Then
        MsgBox "Packet Error.", vbOKOnly
        DestroyGame

        Exit Sub

    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
End Sub

' ::::::::::::::::::::::::::
' :: Alert message packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAlertMsg(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim msg    As String

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
    
    Call MsgBox(msg, vbOKOnly, GAME_NAME)
End Sub

' :::::::::::::::::::::::::::
' :: All characters packet ::
' :::::::::::::::::::::::::::
Private Sub HandleAllChars(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim i      As Long

    Dim Level  As Long

    Dim name   As String

    Dim msg    As String

    Dim Buffer As clsBuffer
    
    ReDim CharAvatars(1 To MAX_CHARS) As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    With frmMainMenu
        .mnuChars.Visible = True
        
        frmSendGetData.Visible = False
        
        .lstChars.Clear
        
        For i = 1 To MAX_CHARS
            CharAvatars(i) = Buffer.ReadLong
            name = Buffer.ReadString
            msg = Buffer.ReadString
            Level = Buffer.ReadByte
            
            If Trim$(name) = vbNullString Then
                .lstChars.AddItem "Free Character Slot"
            Else
                .lstChars.AddItem name & " a level " & Level & " " & msg
            End If

        Next
        
        .lstChars.ListIndex = 0
    End With

End Sub

' :::::::::::::::::::::::::::::::::
' :: Login was successful packet ::
' :::::::::::::::::::::::::::::::::
Private Sub HandleLoginOk(ByVal Index As Long, _
                          ByRef Data() As Byte, _
                          ByVal StartAddr As Long, _
                          ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    frmSendGetData.Visible = True
    frmMainMenu.Visible = False
    frmMainMenu.mnuChars.Visible = False
    
    Call SetStatus("Receiving game data...")

End Sub

' :::::::::::::::::::::::::::::::::::::::
' :: New character classes data packet ::
' :::::::::::::::::::::::::::::::::::::::
Private Sub HandleNewCharClasses(ByVal Index As Long, _
                                 ByRef Data() As Byte, _
                                 ByVal StartAddr As Long, _
                                 ByVal ExtraVar As Long)

    Dim i      As Long

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Max classes
    Max_Classes = Buffer.ReadByte
    ReDim Class(1 To Max_Classes)
    
    For i = 1 To Max_Classes

        With Class(i)
            .name = Buffer.ReadString
            .Avatar = Buffer.ReadLong

            For n = 1 To Vitals.Vital_Count - 1
                .Vital(n) = Buffer.ReadLong
            Next

            For n = 1 To Stats.Stat_Count - 1
                .Stat(n) = Buffer.ReadByte
            Next

        End With

    Next
    
    ' Used for if the player is creating a new character
    With frmMainMenu
        .mnuNewCharacter.Visible = True
        .picPic.Picture = LoadPicture(App.Path & AVATARS_PATH & "1.bmp")
        NewCharAvatar = 1
        
        .picPic.ScaleMode = 3
        .picPic.AutoRedraw = True
        .picPic.PaintPicture .picPic.Picture, _
        0, 0, .picPic.ScaleWidth, .picPic.ScaleHeight, _
        0, 0, .picPic.Picture.Width / 26.46, _
        .picPic.Picture.Height / 26.46
    
        .picPic.Picture = .picPic.Image
        
        frmSendGetData.Visible = False
        
        .cmbClass.Clear
        
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).name)
        Next
        
        .cmbClass.ListIndex = 0
        
        n = .cmbClass.ListIndex + 1
        
        .lblHP.Caption = CStr(Class(n).Vital(Vitals.HP))
        .lblMP.Caption = CStr(Class(n).Vital(Vitals.MP))
        .lblSP.Caption = CStr(Class(n).Vital(Vitals.SP))
        
        .lblSTR.Caption = CStr(Class(n).Stat(Stats.Strength))
        .lblDEF.Caption = CStr(Class(n).Stat(Stats.Defense))
        .lblSPEED.Caption = CStr(Class(n).Stat(Stats.speed))
        .lblMAGI.Caption = CStr(Class(n).Stat(Stats.Magic))
    End With

End Sub

' :::::::::::::::::::::::::
' :: Classes data packet ::
' :::::::::::::::::::::::::
Private Sub HandleClassesData(ByVal Index As Long, _
                              ByRef Data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim n      As Long

    Dim i      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Max classes
    Max_Classes = Buffer.ReadByte
    ReDim Class(1 To Max_Classes)
    
    For i = 1 To Max_Classes

        With Class(i)
            .name = Buffer.ReadString
            .Avatar = Buffer.ReadLong

            For n = 1 To Vitals.Vital_Count - 1
                .Vital(n) = Buffer.ReadLong
            Next

            For n = 1 To Stats.Stat_Count - 1
                .Stat(n) = Buffer.ReadByte
            Next

        End With

    Next

End Sub

' ::::::::::::::::::::
' :: In game packet ::
' ::::::::::::::::::::
Private Sub HandleInGame()
    InGame = True
    Call GameInit
    Call GameLoop
End Sub

' :::::::::::::::::::::::::::::
' :: Player inventory packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerInv(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim n      As Long

    Dim i      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemDur(MyIndex, i, Buffer.ReadLong)
    Next

    Call UpdateInventory
End Sub

' ::::::::::::::::::::::::::::::::::::
' :: Player inventory update packet ::
' ::::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerInvUpdate(ByVal Index As Long, _
                                  ByRef Data() As Byte, _
                                  ByVal StartAddr As Long, _
                                  ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong)
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong)
    Call SetPlayerInvItemDur(MyIndex, n, Buffer.ReadLong)
    Call UpdateInventory
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player worn equipment packet ::
' ::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerWornEq(ByVal Index As Long, _
                               ByRef Data() As Byte, _
                               ByVal StartAddr As Long, _
                               ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim i      As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To Equipment.Equipment_Count - 1
        Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadByte, i)
    Next

    Call UpdateInventory
End Sub

' ::::::::::::::::::::::
' :: Player hp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerHp(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxHP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
    
    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        frmMainGame.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        frmMainGame.picHP.Width = (CLng(GetPlayerVital(MyIndex, Vitals.HP)) / CLng(GetPlayerMaxVital(MyIndex, Vitals.HP))) * 93
    End If

End Sub

' ::::::::::::::::::::::
' :: Player mp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerMp(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxMP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
    
    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        frmMainGame.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        frmMainGame.picMP.Width = (CLng(GetPlayerVital(MyIndex, Vitals.MP)) / CLng(GetPlayerMaxVital(MyIndex, Vitals.MP))) * 93
    End If

End Sub

' ::::::::::::::::::::::
' :: Player sp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerSp(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxSP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.SP, Buffer.ReadLong)
End Sub

' :::::::::::::::::::::::::::
' :: Player stamina packet ::
' :::::::::::::::::::::::::::
Private Sub HandlePlayerStamina(ByVal Index As Long, _
                                ByRef Data() As Byte, _
                                ByVal StartAddr As Long, _
                                ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxStamina = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.Stamina, Buffer.ReadLong)
    
    frmMainGame.lblStamina.Caption = GetPlayerVital(MyIndex, Vitals.Stamina) & "/" & GetPlayerMaxVital(MyIndex, Vitals.Stamina)
    frmMainGame.picStamina.Width = (CLng(GetPlayerVital(MyIndex, Vitals.Stamina)) / CLng(GetPlayerMaxVital(MyIndex, Vitals.Stamina))) * 93
End Sub

' :::::::::::::::::::::::::
' :: Player stats packet ::
' :::::::::::::::::::::::::
Private Sub HandlePlayerStats(ByVal Index As Long, _
                              ByRef Data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim i      As Long, h As Long, B As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        Call SetPlayerStat(Index, i, Buffer.ReadLong)
    Next
    
    Call SetPlayerLevel(Index, Buffer.ReadLong)
    Call SetPlayerName(Index, Buffer.ReadString)
    
    ' Display stats on main UI
    With frmMainGame
        .lblLevel.Caption = GetPlayerLevel(Index)
        .lblSTR.Caption = GetPlayerStat(Index, Stats.Strength)
        .lblDEF.Caption = GetPlayerStat(Index, Stats.Defense)
        .lblMAGI.Caption = GetPlayerStat(Index, Stats.Magic)
        .lblSPEED.Caption = GetPlayerStat(Index, Stats.speed)
        
        h = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
        B = (GetPlayerStat(Index, Stats.Defense) \ 2) + (GetPlayerLevel(Index) \ 2)

        If h > 100 Then h = 100
        If B > 100 Then B = 100
        .lblCritHit.Caption = h & "%"
        .lblBlockChance.Caption = B & "%"
    End With

End Sub

' ::::::::::::::::::::::::
' :: Player data packet ::
' ::::::::::::::::::::::::
Private Sub HandlePlayerData(ByVal Index As Long, _
                             ByRef Data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim i       As Long

    Dim tempRoom As Long

    Dim Buffer  As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    tempRoom = GetPlayerRoom(i)
    
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerAvatar(i, Buffer.ReadLong)
    Call SetPlayerRoom(i, Buffer.ReadLong)
    Player(i).Guild = Buffer.ReadString
    Player(i).GuildAccess = Buffer.ReadLong
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    
    Call GetPlayersInRoom
End Sub

' :::::::::::::::::::::::::::::::::::
' :: Player Exp information packet ::
' :::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerExp(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim NextLevel As Long

    Dim Buffer    As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    SetPlayerExp Index, Buffer.ReadLong
    NextLevel = Buffer.ReadLong
    
    frmMainGame.lblXP.Caption = GetPlayerExp(Index) & "/" & NextLevel
    frmMainGame.picXP.Width = (CLng(GetPlayerExp(Index)) / CLng(NextLevel)) * 273
    
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByVal Index As Long, _
                         ByRef Data() As Byte, _
                         ByVal StartAddr As Long, _
                         ByVal ExtraVar As Long)

    Dim i      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
End Sub

' :::::::::::::::::::::::
' :: NPC attack packet ::
' :::::::::::::::::::::::
Private Sub HandleNpcAttack(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim i      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    ' Set player to attacking
    RoomNpc(i, GetPlayerRoom(MyIndex)).Attacking = 1
    RoomNpc(i, GetPlayerRoom(MyIndex)).AttackTimer = GetTickCount
End Sub

' ::::::::::::::::::::::::::
' :: Check for Room packet ::
' ::::::::::::::::::::::::::
Private Sub HandleCheckForRoom(ByVal Index As Long, _
                              ByRef Data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim x       As Long

    Dim y       As Long

    Dim i       As Long

    Dim NeedRoom As Byte

    Dim Buffer  As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Erase all players except self
    For i = 1 To High_Index

        If i <> MyIndex Then
            Call SetPlayerRoom(i, 0)
        End If

    Next
    
    Call ClearRoomNpcs
    Call ClearRoomItems
    Call ClearRoom
    
    ItemSel = 0
    NPCSel = 0
    PlayerSel = 0
    
    With frmMainGame
        .lstNPCs.ListIndex = -1
        .lstItems.ListIndex = -1
        .lstPlayers.ListIndex = -1
        .picTarget.Picture = LoadPicture(App.Path & AVATARS_PATH & 0 & GFX_EXT)
        .lblTarget.Caption = vbNullString
    End With
    
    ' Get Room num
    x = Buffer.ReadLong
    
    ' Get revision
    y = Buffer.ReadLong
    
    NeedRoom = 1
    
    If FileExist(ROOM_PATH & "room" & x & ROOM_EXT, False) Then
        Call LoadRooms(x)
        
        ' Check to see if the revisions match
        NeedRoom = 1

        If Room.Revision = y Then
            ' We do so we dont need the Room
            'Call SendData(CNeedRoom & SEP_CHAR & "n" & END_CHAR)
            NeedRoom = 0
        End If
    End If
    
    ' Either the revisions didn't match or we dont have the Room, so we need it
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 3
    Buffer.WriteInteger CNeedRoom
    Buffer.WriteByte NeedRoom
    Call SendData(Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Room data packet ::
' :::::::::::::::::::::
Private Sub HandleRoomData(ByVal Index As Long, _
                          ByRef Data() As Byte, _
                          ByVal StartAddr As Long, _
                          ByVal ExtraVar As Long)

    Dim RoomNum    As Long

    Dim Buffer    As clsBuffer

    Dim RoomSize   As Long

    Dim RoomData() As Byte
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    RoomNum = Buffer.ReadLong
    
    If RoomNum < 0 Or RoomNum > MAX_ROOMS Then

        Exit Sub

    End If
    
    RoomSize = LenB(Room)
    ReDim RoomData(RoomSize - 1)
    RoomData = Buffer.ReadBytes(RoomSize)
    CopyMemory ByVal VarPtr(Room), ByVal VarPtr(RoomData(0)), RoomSize
    
    ' Save the Room
    Call SaveRoom(RoomNum)
    
    If InGame Then
        If Player(MyIndex).Room = RoomNum Then
            Call LoadRooms(RoomNum)
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Room items data packet ::
' :::::::::::::::::::::::::::
Private Sub HandleRoomItemData(ByVal Index As Long, _
                              ByRef Data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim i      As Long

    Dim Buffer As clsBuffer

    Dim RoomNum As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    RoomNum = Buffer.ReadLong
    
    For i = 1 To MAX_ROOM_ITEMS

        With RoomItem(i, RoomNum)
            .num = Buffer.ReadByte
            .Value = Buffer.ReadLong
            .Dur = Buffer.ReadInteger
        End With

    Next

End Sub

' :::::::::::::::::::::::::
' :: Room npc data packet ::
' :::::::::::::::::::::::::
Private Sub HandleRoomNpcData(ByVal Index As Long, _
                             ByRef Data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim i      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Dim Room As Long

    Room = Buffer.ReadLong
    
    For i = 1 To MAX_ROOM_NPCS

        With RoomNpc(i, Room)
            .num = Buffer.ReadInteger
        End With

    Next

End Sub

' :::::::::::::::::::::::::::::::
' :: Room send completed packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRoomDone()

    Dim i         As Long

    Dim MusicFile As String
    
    MusicFile = Trim$(CStr(Room.Music))
    
    ' Get high NPC index
    High_Npc_Index = 0

    For i = 1 To MAX_ROOM_NPCS

        If Room.Npc(i) > 0 Then
            High_Npc_Index = High_Npc_Index + 1
        Else

            Exit For

        End If

    Next
    
    ' Play music
    If GameData.Music = 1 Then
        If Len(Room.Music) > 0 Then
            If MusicFile <> CurrentMusic Then
                Stop_Music
                Play_Music (MusicFile)
                CurrentMusic = MusicFile
            End If
    
        Else
            Stop_Music
            CurrentMusic = 0
        End If
    End If
    
    If Room.Shop = 0 Then
        frmMainGame.picTradeButton.Visible = False
    Else
        frmMainGame.picTradeButton.Visible = True
    End If
    
    Call RefreshEntityList(MyIndex, 1)
    Call RefreshEntityList(MyIndex, 2)
    
    GettingRoom = False
    CanMoveNow = True
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, _
                         ByRef Data() As Byte, _
                         ByVal StartAddr As Long, _
                         ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, _
                               ByRef Data() As Byte, _
                               ByVal StartAddr As Long, _
                               ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

Private Sub HandleRoomMsg(ByVal Index As Long, _
                         ByRef Data() As Byte, _
                         ByVal StartAddr As Long, _
                         ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim msg    As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    Call AddText(msg)
End Sub

' :::::::::::::::::::::::::::
' :: Refresh editor packet ::
' :::::::::::::::::::::::::::
Private Sub HandleRefresh()

    Dim i As Long
    
    frmIndex.lstIndex.Clear
    
    Select Case Editor

        Case EDITOR_ITEM

            For i = 1 To MAX_ITEMS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).name)
            Next

        Case EDITOR_NPC

            For i = 1 To MAX_NPCS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).name)
            Next

        Case EDITOR_SHOP

            For i = 1 To MAX_SHOPS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
            Next

        Case EDITOR_SPELL

            For i = 1 To MAX_SPELLS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).name)
            Next

    End Select
    
    frmIndex.lstIndex.ListIndex = 0
    
End Sub

' :::::::::::::::::::::::
' :: Item spawn packet ::
' :::::::::::::::::::::::
Private Sub HandleSpawnItem(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer

    Dim RoomNum As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    RoomNum = Buffer.ReadLong
    n = Buffer.ReadLong
    
    With RoomItem(n, RoomNum)
        .num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .Dur = Buffer.ReadLong
    End With
    
    Call RefreshEntityList(MyIndex, 2)
End Sub

' ::::::::::::::::::::::::
' :: Item editor packet ::
' ::::::::::::::::::::::::
Private Sub HandleItemEditor()

    Dim i As Long
    
    With frmIndex
        .Caption = "Item Index"
        Editor = EDITOR_ITEM
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With

End Sub

' ::::::::::::::::::::::::
' :: Update item packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateItem(ByVal Index As Long, _
                             ByRef Data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the item
    With Item(n)
        .name = Buffer.ReadString
        .Pic = Buffer.ReadInteger
        .Type = Buffer.ReadByte
        .Data1 = Buffer.ReadInteger
        .Data2 = Buffer.ReadInteger
        .Data3 = Buffer.ReadInteger
    End With
    
    Call RefreshEntityList(MyIndex, 2)
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Private Sub HandleEditItem(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim ItemNum    As Long

    Dim ItemSize   As Long

    Dim ItemData() As Byte

    Dim Buffer     As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then

        Exit Sub

    End If
    
    ' Update the item
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
    
    ' Initialize the item editor
    Call ItemEditorInit
End Sub

' ::::::::::::::::::::::
' :: Npc spawn packet ::
' ::::::::::::::::::::::
Private Sub HandleSpawnNpc(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim n, Room As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Room = Buffer.ReadLong
    n = Buffer.ReadLong
    
    With RoomNpc(n, Room)
        .num = Buffer.ReadInteger
    End With
    
    Call RefreshEntityList(Index, 1)
        
End Sub

' :::::::::::::::::::::
' :: Npc dead packet ::
' :::::::::::::::::::::
Private Sub HandleNpcDead(ByVal Index As Long, _
                          ByRef Data() As Byte, _
                          ByVal StartAddr As Long, _
                          ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer

    Dim tRoom   As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    tRoom = Buffer.ReadLong
    n = Buffer.ReadLong
    Call ClearRoomNpc(n, tRoom)
    
    Call RefreshEntityList(Index, 1)
    
End Sub

' :::::::::::::::::::::::
' :: Npc editor packet ::
' :::::::::::::::::::::::
Private Sub HandleNpcEditor()

    Dim i As Long
    
    With frmIndex
        .Caption = "NPC Index"
        Editor = EDITOR_NPC
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With

End Sub

' :::::::::::::::::::::::
' :: Update npc packet ::
' :::::::::::::::::::::::
Private Sub HandleUpdateNpc(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the NPC
    With Npc(n)
        .name = Buffer.ReadString
        .AttackSay = vbNullString
        .Avatar = Buffer.ReadInteger
        .SpawnSecs = 0
        .Behavior = 0
        .Range = 0
        .DropChance = 0
        .DropItem = 0
        .DropItemValue = 0
        .Stat(Stats.Strength) = 0
        .Stat(Stats.Defense) = 0
        .Stat(Stats.speed) = 0
        .Stat(Stats.Magic) = 0
    End With
    
    Call RefreshEntityList(Index, 1)
End Sub

' :::::::::::::::::::::
' :: Edit npc packet ::
' :::::::::::::::::::::
Private Sub HandleEditNpc(ByVal Index As Long, _
                          ByRef Data() As Byte, _
                          ByVal StartAddr As Long, _
                          ByVal ExtraVar As Long)

    Dim NPCNum    As Long

    Dim NpcSize   As Long

    Dim NpcData() As Byte

    Dim Buffer    As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    NPCNum = Buffer.ReadLong
    
    If NPCNum < 0 Or NPCNum > MAX_NPCS Then

        Exit Sub

    End If
    
    ' Update the Npc
    NpcSize = LenB(Npc(NPCNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(NPCNum)), ByVal VarPtr(NpcData(0)), NpcSize
    
    ' Initialize the npc editor
    Call NpcEditorInit
End Sub

' :::::::::::::::::::::
' :: Edit Room packet ::
' :::::::::::::::::::::
Private Sub HandleEditRoom()
    Call RoomEditorInit
End Sub

' ::::::::::::::::::::::::
' :: Shop editor packet ::
' ::::::::::::::::::::::::
Private Sub HandleShopEditor()

    Dim i As Long
    
    With frmIndex
        .Caption = "Shop Index"
        Editor = EDITOR_SHOP
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With

End Sub

' ::::::::::::::::::::::::
' :: Update shop packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateShop(ByVal Index As Long, _
                             ByRef Data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the shop name
    Shop(n).name = Buffer.ReadString
End Sub

' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Private Sub HandleEditShop(ByVal Index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim ShopNum    As Long

    Dim ShopSize   As Long

    Dim ShopData() As Byte

    Dim Buffer     As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then

        Exit Sub

    End If
    
    ' Update the Shop
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    ' Initialize the shop editor
    Call ShopEditorInit
    
End Sub

' :::::::::::::::::::::::::
' :: Spell editor packet ::
' :::::::::::::::::::::::::
Private Sub HandleSpellEditor()

    Dim i As Long
    
    With frmIndex
        .Caption = "Spell Index"
        Editor = EDITOR_SPELL
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With

End Sub

' ::::::::::::::::::::::::
' :: Update spell packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateSpell(ByVal Index As Long, _
                              ByRef Data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim n      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the spell name
    With Spell(n)
        .name = Buffer.ReadString
        .MPReq = Buffer.ReadInteger
        .Pic = Buffer.ReadInteger
    End With

End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
Private Sub HandleEditSpell(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim spellnum    As Long

    Dim SpellSize   As Long

    Dim SpellData() As Byte

    Dim Buffer      As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
    
    If spellnum < 0 Or spellnum > MAX_SPELLS Then

        Exit Sub

    End If
    
    ' Update the Spell
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    
    ' Initialize the spell editor
    Call SpellEditorInit
End Sub

' ::::::::::::::::::
' :: Trade packet ::
' ::::::::::::::::::
Private Sub HandleTrade(ByVal Index As Long, _
                        ByRef Data() As Byte, _
                        ByVal StartAddr As Long, _
                        ByVal ExtraVar As Long)

    Dim i         As Long

    Dim ShopNum   As Long

    Dim GiveItem  As Long

    Dim GiveValue As Long

    Dim GetItem   As Long

    Dim GetValue  As Long

    Dim Buffer    As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    With frmTrade

        If Buffer.ReadByte = 1 Then
            .lblFixItem.Visible = True
        Else
            .lblFixItem.Visible = False
        End If
        
        For i = 1 To MAX_TRADES
            GiveItem = Buffer.ReadLong
            GiveValue = Buffer.ReadLong
            GetItem = Buffer.ReadLong
            GetValue = Buffer.ReadLong
            
            If GiveItem > 0 Then
                If GetItem > 0 Then
                    .lstTrade.AddItem "Give " & Trim$(Shop(ShopNum).name) & " " & GiveValue & " " & Trim$(Item(GiveItem).name) & " for " & GetValue & " " & Trim$(Item(GetItem).name)
                End If
            End If

        Next
        
        If .lstTrade.ListCount > 0 Then
            .lstTrade.ListIndex = 0
        End If

        .Show vbModal
    End With

End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Private Sub HandleSpells(ByVal Index As Long, _
                         ByRef Data() As Byte, _
                         ByVal StartAddr As Long, _
                         ByVal ExtraVar As Long)

    Dim i      As Long

    Dim j      As Long

    Dim n      As Long

    Dim k      As Long

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    With frmMainGame
        .lstSpells.Clear
        
        For i = 1 To MAX_PLAYER_SPELLS
            k = Buffer.ReadLong
            PlayerSpells(k) = Buffer.ReadLong
        Next
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS

            If PlayerSpells(i) <> 0 Then
                .lstSpells.AddItem i & ": " & Trim$(Spell(PlayerSpells(i)).name)
            Else
                .lstSpells.AddItem " "
            End If

        Next
        
        .lstSpells.ListIndex = 0
    End With

End Sub

' ::::::::::::::::::::::
' :: Left game packet ::
' ::::::::::::::::::::::
Private Sub HandleLeft(ByVal Index As Long, _
                       ByRef Data() As Byte, _
                       ByVal StartAddr As Long, _
                       ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Call ClearPlayer(Buffer.ReadLong)
    Call GetPlayersInRoom
End Sub

' ::::::::::::::::::::::
' :: HighIndex packet ::
' ::::::::::::::::::::::
Private Sub HandleHighIndex(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    High_Index = Buffer.ReadLong
End Sub

' :::::::::::::::::::::::
' :: Spell Cast packet ::
' :::::::::::::::::::::::
Private Sub HandleSpellCast(ByVal Index As Long, _
                            ByRef Data() As Byte, _
                            ByVal StartAddr As Long, _
                            ByVal ExtraVar As Long)

    Dim i          As Long

    Dim TargetType As Byte

    Dim n          As Long

    Dim spellnum   As Long

    Dim Buffer     As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    TargetType = Buffer.ReadByte
    n = Buffer.ReadLong
    spellnum = Buffer.ReadLong
    
    If n = 0 Or spellnum = 0 Then

        Exit Sub

    End If
    
End Sub

Private Sub HandleDoor(ByVal Index As Long, _
                       ByRef Data() As Byte, _
                       ByVal StartAddr As Long, _
                       ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
End Sub

Private Sub HandleMaxes(ByVal Index As Long, _
                        ByRef Data() As Byte, _
                        ByVal StartAddr As Long, _
                        ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MAX_PLAYERS = Buffer.ReadInteger
    MAX_ITEMS = Buffer.ReadInteger
    MAX_NPCS = Buffer.ReadInteger
    MAX_SHOPS = Buffer.ReadInteger
    MAX_SPELLS = Buffer.ReadInteger
    MAX_ROOMS = Buffer.ReadInteger
    
    ReDim RoomNpc(1 To MAX_ROOM_NPCS, 1 To MAX_ROOMS)
    ReDim RoomItem(1 To MAX_ROOM_ITEMS, 1 To MAX_ROOMS)
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim PlayerLst(1 To MAX_PLAYERS) As Long
    
End Sub

Private Sub HandleSync(ByVal Index As Long, _
                       ByRef Data() As Byte, _
                       ByVal StartAddr As Long, _
                       ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer

    Dim tm As Long

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()

    tm = Buffer.ReadLong

    If SyncRoom = tm Then
        SentSync = False

        Exit Sub

    End If
    Player(MyIndex).Room = tm

    SentSync = False
End Sub

Private Sub HandleRoomRevs(ByVal Index As Long, _
                          ByRef Data() As Byte, _
                          ByVal StartAddr As Long, _
                          ByVal ExtraVar As Long)

    Dim Buffer  As clsBuffer

    Dim i       As Long

    Dim Buffer2 As clsBuffer
    
    Set Buffer2 = New clsBuffer
    
    Buffer2.PreAllocate MAX_ROOMS * 1 + 2
    
    Buffer2.WriteInteger CRoomReqs
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_ROOMS

        'If CheckRoomRevision(i, Buffer.ReadLong) = False Then
        '    Buffer2.WriteByte 1
        'Else
            Buffer2.WriteByte 0
        'End If

    Next
    
    Call SendData(Buffer2.ToArray())
End Sub

