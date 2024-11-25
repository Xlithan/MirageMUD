Attribute VB_Name = "modServerTCP"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Sub UpdateCaption()
    frmServer.Caption = "MirageMUD Server <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalPlayersOnline & ")"
End Sub

Public Sub CreateFullRoomCache()
    Dim i As Long
    
    For i = 1 To MAX_ROOMS
        Call RoomCache_Create(i)
    Next
    
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If LCase$(Trim$(Player(PlayersOnline(i)).Login)) = LCase$(Login) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next
    
End Function

Function IsMultiChars(ByVal Name As String) As Boolean
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If LCase$(GetPlayerName(i)) = LCase$(Name) Then
            IsMultiChars = True
            Exit Function
        End If
    Next
    
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long
    
    For i = 1 To TotalPlayersOnline
        If Trim$(GetPlayerIP(PlayersOnline(i))) = IP Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
            
        End If
    Next
End Function

Private Function IsBanned(ByVal IP As String) As Boolean
    Dim FileName As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    
    FileName = App.Path & "\data\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    F = FreeFile
    Open FileName For Input As #F
    
    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName
        
        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If
    Loop
    
    Close #F
End Function

Public Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteInteger (UBound(Data) - LBound(Data)) + 1 ' Writes the length of the packet
        Buffer.WriteBytes Data()            ' Writes the data to the packet
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If
End Sub

Public Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        Call SendDataTo(PlayersOnline(i), Data)
    Next
End Sub

Public Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            Call SendDataTo(PlayersOnline(i), Data)
        End If
    Next
End Sub

Public Sub SendDataToRoom(ByVal RoomNum As Long, ByRef Data() As Byte)
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerRoom(PlayersOnline(i)) = RoomNum Then
            Call SendDataTo(PlayersOnline(i), Data)
        End If
    Next
End Sub

Public Sub SendDataToRoomBut(ByVal Index As Long, ByVal RoomNum As Long, ByRef Data() As Byte)
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerRoom(PlayersOnline(i)) = RoomNum Then
            If PlayersOnline(i) <> Index Then
                Call SendDataTo(PlayersOnline(i), Data)
            End If
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SGlobalMsg
    Buffer.WriteString Msg
    
    Call SendDataToAll(Buffer.ToArray)
End Sub

Public Sub AdminMsg(ByVal Msg As String)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SAdminMsg
    Buffer.WriteString Msg
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerAccess(PlayersOnline(i)) > 0 Then
            Call SendDataTo(PlayersOnline(i), Buffer.ToArray)
        End If
    Next
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SPlayerMsg
    Buffer.WriteString Msg
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub RoomMsg(ByVal RoomNum As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SRoomMsg
    Buffer.WriteString Msg
    
    Call SendDataToRoom(RoomNum, Buffer.ToArray)
End Sub

Public Sub RoomMsgBut(ByVal RoomNum As Long, ByVal pID As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SRoomMsg
    Buffer.WriteString Msg
    
    Call SendDataToRoomBut(pID, RoomNum, Buffer.ToArray)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 4
    Buffer.WriteInteger SAlertMsg
    Buffer.WriteString Msg
    
    Call SendDataTo(Index, Buffer.ToArray)
    DoEvents
    Call CloseSocket(Index)
End Sub

Public Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(COLOR_WHITE & GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")")
        End If
        
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Public Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long
    
    If (Index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Public Sub SocketConnected(ByVal Index As Long)
    ' Are they trying to connect more then one connection?
    'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
    If Not IsBanned(GetPlayerIP(Index)) Then
        Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
    Else
        Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
    End If
    'Else
    ' Tried multiple connections
    '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
    'End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long
    
    ' Get the data as an array
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    
    ' Write the bytes to the byte array
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    ' Check if we have enough in the buffer
    If TempPlayer(Index).Buffer.Length >= 2 Then
        pLength = TempPlayer(Index).Buffer.ReadInteger(False)
        
        ' If the plength is less than 0 then we know there was something odd
        ' hacking attempt is usually what happened
        If pLength < 0 Then
            HackingAttempt Index, "Hacking attempt."
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 2
        If pLength <= TempPlayer(Index).Buffer.Length - 2 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            
            ' Remove the "size" off the packet now that we have the full packet
            TempPlayer(Index).Buffer.ReadInteger
            ' Handle the packet data
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 2 Then
            pLength = TempPlayer(Index).Buffer.ReadInteger(False)
            
            If pLength < 0 Then
                HackingAttempt Index, "Hacking attempt."
                Exit Sub
            End If
        End If
    Loop
    
    ' Trim down the packet
    TempPlayer(Index).Buffer.Trim
    
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    
    If GetTickCount >= TempPlayer(Index).DataTimer + 1000 Then
        TempPlayer(Index).DataTimer = GetTickCount
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
        Exit Sub
    End If
    
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 2000 Then
            Call HackingAttempt(Index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 55 Then
            Call HackingAttempt(Index, "Packet Flooding")
            Exit Sub
        End If
    End If
    
End Sub

Public Sub CloseSocket(ByVal Index As Long)
    
    If Index > 0 Then
        Call LeftGame(Index)
        
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        
        frmServer.Socket(Index).Close
        
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Public Sub RoomCache_Create(ByVal RoomNum As Long)
    Dim RoomSize As Long
    Dim RoomData() As Byte
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    RoomSize = LenB(Room(RoomNum))
    ReDim RoomData(RoomSize - 1)
    CopyMemory RoomData(0), ByVal VarPtr(Room(RoomNum)), RoomSize
    
    Buffer.PreAllocate RoomSize + 6
    Buffer.WriteInteger SRoomData
    Buffer.WriteLong RoomNum
    Buffer.WriteBytes RoomData
    
    RoomCache(RoomNum).Cache = Buffer.ToArray()
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************

Public Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            s = s & GetPlayerName(PlayersOnline(i)) & ", "
            n = n + 1
        End If
    Next
    
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
    
    Call PlayerMsg(Index, COLOR_WHO & s)
End Sub

Public Sub SendChars(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SAllChars
    For i = 1 To MAX_CHARS
        Buffer.WriteLong Player(Index).Char(i).Avatar
        Buffer.WriteString Trim$(Player(Index).Char(i).Name)
        Buffer.WriteString Trim$(Class(Player(Index).Char(i).Class).Name)
        Buffer.WriteByte Player(Index).Char(i).Level
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendMaxes(ByVal Index As Long)
    Dim Buffer As New clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 14
    Buffer.WriteInteger SSendMaxes
    Buffer.WriteInteger MAX_PLAYERS
    Buffer.WriteInteger MAX_ITEMS
    Buffer.WriteInteger MAX_NPCS
    Buffer.WriteInteger MAX_SHOPS
    Buffer.WriteInteger MAX_SPELLS
    Buffer.WriteInteger MAX_ROOMS
    
    Call SendDataTo(Index, Buffer.ToArray)
    
End Sub

Public Sub SendJoinRoom(ByVal Index As Long)
    Dim i As Long
    ' Send all players on current Room to index
    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            'If GetPlayerRoom(PlayersOnline(i)) = GetPlayerRoom(Index) Then
            Call SendDataTo(Index, PlayerData(PlayersOnline(i)))
            'End If
        End If
    Next
    
    ' Send index's player data to everyone on the Room including himself
    Call SendDataToAll(PlayerData(Index))
End Sub

Public Sub SendLeaveRoom(ByVal Index As Long, ByVal RoomNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger SLeft
    Buffer.WriteLong Index
    
    Call SendDataToRoomBut(Index, RoomNum, Buffer.ToArray())
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
    ' Send index's player data to everyone including himself on the Room
    Call SendDataToRoom(GetPlayerRoom(Index), PlayerData(Index))
End Sub

Public Sub SendRoom(ByVal Index As Long, ByVal RoomNum As Long)
    Call SendDataTo(Index, RoomCache(RoomNum).Cache)
End Sub

Public Sub SendRoomItemsTo(ByVal Index As Long, ByVal RoomNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_ROOM_ITEMS * 9) + 2 + 4
    Buffer.WriteInteger SRoomItemData
    Buffer.WriteLong RoomNum
    For i = 1 To MAX_ROOM_ITEMS
        Buffer.WriteByte RoomItem(RoomNum, i).Num
        Buffer.WriteLong RoomItem(RoomNum, i).Value
        Buffer.WriteInteger RoomItem(RoomNum, i).Dur
        Buffer.WriteByte RoomItem(RoomNum, i).X
        Buffer.WriteByte RoomItem(RoomNum, i).y
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendRoomItemsToAll(ByVal RoomNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_ROOM_ITEMS * 9) + 2 + 4
    Buffer.WriteInteger SRoomItemData
    Buffer.WriteLong RoomNum
    For i = 1 To MAX_ROOM_ITEMS
        Buffer.WriteByte RoomItem(RoomNum, i).Num
        Buffer.WriteLong RoomItem(RoomNum, i).Value
        Buffer.WriteInteger RoomItem(RoomNum, i).Dur
        Buffer.WriteByte RoomItem(RoomNum, i).X
        Buffer.WriteByte RoomItem(RoomNum, i).y
    Next
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendRoomNpcsTo(ByVal Index As Long, ByVal RoomNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_ROOM_NPCS * 6) + 6
    Buffer.WriteInteger SRoomNpcData
    Buffer.WriteLong RoomNum
    For i = 1 To MAX_ROOM_NPCS
        Buffer.WriteInteger RoomNpc(RoomNum, i).Num
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendRoomNpcsToRoom(ByVal RoomNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_ROOM_NPCS * 6) + 6
    Buffer.WriteInteger SRoomNpcData
    Buffer.WriteLong RoomNum
    For i = 1 To MAX_ROOM_NPCS
        Buffer.WriteInteger RoomNpc(RoomNum, i).Num
    Next
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendRoomRevs(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_ROOMS * 4) + 2
    Buffer.WriteInteger SRoomRevs
    For i = 1 To MAX_ROOMS
        Buffer.WriteLong Room(i).Revision
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendItems(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_ITEMS
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next
End Sub

Public Sub SendNpcs(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_NPCS
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next
    
    For i = 1 To MAX_ROOMS
        Call SendRoomNpcsTo(Index, i)
    Next
End Sub

Public Sub SendInventory(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_INV * 12) + 2
    Buffer.WriteInteger SPlayerInv
    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteLong GetPlayerInvItemDur(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 18
    Buffer.WriteInteger SPlayerInvUpdate
    Buffer.WriteLong InvSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemDur(Index, InvSlot)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (Equipment.Equipment_Count - 1) + 2
    Buffer.WriteInteger SPlayerWornEq
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteByte GetPlayerEquipmentSlot(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Select Case Vital
    Case HP
        Buffer.WriteInteger SPlayerHp
    Case MP
        Buffer.WriteInteger SPlayerMp
    Case SP
        Buffer.WriteInteger SPlayerSp
    Case Stamina
        Buffer.WriteInteger SPlayerStamina
    End Select
    
    Buffer.WriteLong GetPlayerMaxVital(Index, Vital)
    Buffer.WriteLong GetPlayerVital(Index, Vital)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendStats(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate ((Stats.Stat_Count - 1) * 4) + 2
    Buffer.WriteInteger SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteString GetPlayerName(Index)
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendWelcome(ByVal Index As Long)
    ' Send them welcome
    Call PlayerMsg(Index, "[font=Courier New][b]                                                [color=#964B00]|[color=#FFFF00]>>>")
    Call PlayerMsg(Index, "[font=Courier New][b]                                                [color=#964B00]|")
    Call PlayerMsg(Index, "[font=Courier New][b]                                            [color=#808080]_  _[color=#964B00]|[color=#808080]_ _")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]___  ____                                 [color=#808080] |;|_|;|_|;|")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]|  \/  (_)                                [color=#808080] \\.    .  /")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| .  . |_ _ __ __ _  __ _  ___            [color=#808080]  \\:  .  /")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| |\/| | | '__/ _` |/ _` |/ _ \           [color=#808080]   ||:   |")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| |  | | | | | (_| | (_| |  __/           [color=#808080]   ||:.  |")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]\_|  |_|_|_|  \__,_|\__, |\___|           [color=#808080]   ||:  .|")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]___  ____   _______  __/ |                [color=#808080]   ||:   |       [color=#FFFFFF]\,/")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]|  \/  | | | |  _  \|___/                 [color=#808080]   ||: , |            [color=#FFFFFF]/`\")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| .  . | | | | | | |                      [color=#808080]   ||:   |")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| |\/| | | | | | | |                      [color=#808080]   ||: . |")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]| |  | | |_| | |/ /                         [color=#66ff00]_[color=#808080]||[color=#66ff00]_   [color=#808080]|")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#0096FF]\_|  |_/\___/|___/[color=#66ff00]--~~__            __ ----~    ~`---,")
    Call PlayerMsg(Index, "[font=Courier New][b][color=#66ff00]-~--~                   ~---__ ,--~'                  ~~----_____[/color][/b][/font]")
    Call PlayerMsg(Index, COLOR_CYAN & "[b]Welcome to the test client for MirageMUD.[/b]")
    
    ' Send them MOTD
    If LenB(MOTD) > 0 Then
        Call PlayerMsg(Index, COLOR_BRIGHTCYAN & "MOTD: " & MOTD)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Public Sub SendClasses(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SClassesData
    Buffer.WriteByte Max_Classes
    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong Class(i).Avatar
        For n = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetClassMaxVital(i, n)
        Next
        For n = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte Class(i).Stat(n)
        Next
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendNewCharClasses(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SNewCharClasses
    Buffer.WriteByte Max_Classes
    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong Class(i).Avatar
        For n = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetClassMaxVital(i, n)
        Next
        For n = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte Class(i).Stat(n)
        Next
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendLeftGame(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger SLeft
    Buffer.WriteLong Index
    
    Call SendDataToAllBut(Index, Buffer.ToArray())
End Sub

Public Sub SendPlayerExp(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteInteger SPlayerExp
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Item(ItemNum)) + 2
    Buffer.WriteInteger SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteInteger Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteInteger Item(ItemNum).Data1
    Buffer.WriteInteger Item(ItemNum).Data2
    Buffer.WriteInteger Item(ItemNum).Data3
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Item(ItemNum)) + 2
    Buffer.WriteInteger SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteInteger Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteInteger Item(ItemNum).Data1
    Buffer.WriteInteger Item(ItemNum).Data2
    Buffer.WriteInteger Item(ItemNum).Data3
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemData() As Byte
    Dim ItemSize As Long
    
    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    
    Buffer.PreAllocate ItemSize + 6
    Buffer.WriteInteger SEditItem
    Buffer.WriteLong ItemNum
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteBytes ItemData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Npc(NpcNum).Name)) + 8
    Buffer.WriteInteger SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Avatar
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Npc(NpcNum).Name)) + 8
    Buffer.WriteInteger SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Avatar
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcData() As Byte
    Dim NpcSize As Long
    
    Set Buffer = New clsBuffer
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize)
    
    Buffer.PreAllocate NpcSize + 6
    Buffer.WriteInteger SEditNpc
    Buffer.WriteLong NpcNum
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Buffer.WriteBytes NpcData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendShops(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_SHOPS
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next
End Sub

Public Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Shop(ShopNum).Name)) + 8
    Buffer.WriteInteger SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Shop(ShopNum).Name)) + 8
    Buffer.WriteInteger SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopData() As Byte
    Dim ShopSize As Long
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize)
    
    Buffer.PreAllocate ShopSize + 6
    Buffer.WriteInteger SEditShop
    Buffer.WriteLong ShopNum
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Buffer.WriteBytes ShopData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendSpells(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_SPELLS
        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next
End Sub

Public Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Spell(SpellNum).Name)) + 12
    Buffer.WriteInteger SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteInteger Spell(SpellNum).MPReq
    Buffer.WriteInteger Spell(SpellNum).Pic
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Spell(SpellNum).Name)) + 12
    Buffer.WriteInteger SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteInteger Spell(SpellNum).MPReq
    Buffer.WriteInteger Spell(SpellNum).Pic
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Dim SpellData() As Byte
    Dim SpellSize As Long
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize)
    
    Buffer.PreAllocate SpellSize + 6
    Buffer.WriteInteger SEditSpell
    Buffer.WriteLong SpellNum
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    Buffer.WriteBytes SpellData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
    Dim i As Long
    Dim X As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_TRADES * 16) + 7
    Buffer.WriteInteger STrade
    Buffer.WriteLong ShopNum
    Buffer.WriteByte Shop(ShopNum).FixesItems
    
    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetValue
        
        ' Item #
        X = Shop(ShopNum).TradeItem(i).GetItem
        
        If X > 0 And X <= MAX_ITEMS Then
            
            If Item(X).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                y = Spell(Item(X).Data1).ClassReq
                
                If y = 0 Then
                    Call PlayerMsg(Index, COLOR_YELLOW & Trim$(Item(X).Name) & " can be used by all classes.")
                Else
                    Call PlayerMsg(Index, COLOR_YELLOW & Trim$(Item(X).Name) & " can only be used by a " & GetClassName(y - 1) & ".")
                End If
            End If
            
        End If
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendPlayerSpells(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_PLAYER_SPELLS * 8) + 2
    Buffer.WriteInteger SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong i
        Buffer.WriteLong GetPlayerSpell(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteInteger SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAvatar(Index)
    Buffer.WriteLong GetPlayerRoom(Index)
    Buffer.WriteString GetPlayerGuild(Index)
    Buffer.WriteLong GetPlayerGAccess(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    PlayerData = Buffer.ToArray()
End Function
