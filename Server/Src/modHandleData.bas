Attribute VB_Name = "modHandleData"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CGetClasses) = GetAddress(AddressOf HandleGetClasses)
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CDelChar) = GetAddress(AddressOf HandleDelChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(CAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetAvatar) = GetAddress(AddressOf HandleSetAvatar)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewRoom) = GetAddress(AddressOf HandleRequestNewRoom)
    HandleDataSub(CRoomData) = GetAddress(AddressOf HandleRoomData)
    HandleDataSub(CNeedRoom) = GetAddress(AddressOf HandleNeedRoom)
    HandleDataSub(CRoomGetItem) = GetAddress(AddressOf HandleRoomGetItem)
    HandleDataSub(CRoomDropItem) = GetAddress(AddressOf HandleRoomDropItem)
    HandleDataSub(CRoomRespawn) = GetAddress(AddressOf HandleRoomRespawn)
    HandleDataSub(CRoomReport) = GetAddress(AddressOf HandleRoomReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditRoom) = GetAddress(AddressOf HandleRequestEditRoom)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CDelete) = GetAddress(AddressOf HandleDelete)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditSpell)
    HandleDataSub(CEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CTrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CFixItem) = GetAddress(AddressOf HandleFixItem)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CParty) = GetAddress(AddressOf HandleParty)
    HandleDataSub(CJoinParty) = GetAddress(AddressOf HandleJoinParty)
    HandleDataSub(CLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSync) = GetAddress(AddressOf HandleSync)
    HandleDataSub(CRoomReqs) = GetAddress(AddressOf HandleRoomReqs)
    HandleDataSub(CSleepinn) = GetAddress(AddressOf HandleSleepInn)
    HandleDataSub(CCreateGuild) = GetAddress(AddressOf HandleCreateGuild)
    HandleDataSub(CRemoveFromGuild) = GetAddress(AddressOf HandleRemoveFromGuild)
    HandleDataSub(CInviteGuild) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(CKickGuild) = GetAddress(AddressOf HandleGuildKick)
    HandleDataSub(CGuildPromote) = GetAddress(AddressOf HandleGuildPromote)
    HandleDataSub(CLeaveGuild) = GetAddress(AddressOf HandleLeaveGuild)
End Sub

' Will handle the packet data
Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msgtype As Integer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msgtype = Buffer.ReadInteger
    
    If msgtype < 0 Or msgtype >= CMSG_COUNT Then
        HackingAttempt Index, "Packet Manipulation."
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(msgtype), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Requesting classes for making a character ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleGetClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then
        Call SendNewCharClasses(Index)
    End If
End Sub

' ::::::::::::::::::::::::
' :: New account packet ::
' ::::::::::::::::::::::::
Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long
    
    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))
                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(Index, "Your account has been created!")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
        End If
    End If
    
End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    
    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
            
            ' Delete names from master name file
            Call LoadPlayer(Index, Name)
            For i = 1 To MAX_CHARS
                If LenB(Trim$(Player(Index).Char(i).Name)) > 0 Then
                    Call DeleteName(Player(Index).Char(i).Name)
                End If
            Next
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call Kill(App.Path & "\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
        End If
    End If
    
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long
    
    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
            ' Check versions
            If Buffer.ReadByte < CLIENT_MAJOR Or Buffer.ReadByte < CLIENT_MINOR Or Buffer.ReadByte < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & GAME_WEBSITE)
                Exit Sub
            End If
            
            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If
            
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
            
            'If IsMultiAccounts(Name) Then
            '    Call AlertMsg(Index, "Multiple account logins is not authorized.")
            '    Exit Sub
            'End If
            
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
            Call SendMaxes(Index)
            Call SendRoomRevs(Index)
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
        End If
    End If
    
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim CharNum As Long
    Dim Avatar As Long
    Dim i As Long
    Dim n As Long
    
    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        
        Buffer.WriteBytes Data()
        
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        CharNum = Buffer.ReadLong
        Avatar = Buffer.ReadLong
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If
        Next
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Call HackingAttempt(Index, "Invalid Sex")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Call HackingAttempt(Index, "Invalid Class")
            Exit Sub
        End If
        
        ' Check if char already exists in slot
        If CharExist(Index, CharNum) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If
        
        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If
        
        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, CharNum, Avatar)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been created!")
    End If
End Sub

' :::::::::::::::::::::::::::::::
' :: Deleting character packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleDelChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim CharNum As Long
    
    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        
        Buffer.WriteBytes Data()
        
        CharNum = Buffer.ReadLong
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        Call DelChar(Index, CharNum)
        Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been deleted!")
    End If
End Sub

' ::::::::::::::::::::::::::::
' :: Using character packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleUseChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim CharNum As Long
    Dim F As Long
    Dim Buffer As clsBuffer
    
    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        
        Buffer.WriteBytes Data()
        
        CharNum = Buffer.ReadLong
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index, CharNum) Then
            TempPlayer(Index).CharNum = CharNum
            
            If IsMultiChars(GetPlayerName(Index)) Then
                Call AlertMsg(Index, "This character is already logged in.")
                Exit Sub
            End If
            
            Call JoinGame(Index)
            
            CharNum = TempPlayer(Index).CharNum
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
            Call UpdateCaption
            
            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
            If Not FindChar(GetPlayerName(Index)) Then
                F = FreeFile
                Open App.Path & "\accounts\charlist.txt" For Append As #F
                Print #F, GetPlayerName(Index)
                Close #F
            End If
        Else
            Call AlertMsg(Index, "Character does not exist!")
        End If
    End If
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Say Text Modification")
            Exit Sub
        End If
    Next
    
    Call AddLog("Room #" & GetPlayerRoom(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call RoomMsg(GetPlayerRoom(Index), COLOR_BRIGHTRED & "[b]" & GetPlayerName(Index) & ": " & COLOR_WHITE & Msg & "[/b]")
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Emote Text Modification")
            Exit Sub
        End If
    Next
    
    Call AddLog("Room #" & GetPlayerRoom(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call RoomMsg(GetPlayerRoom(Index), COLOR_EMOTE & GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1))
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Broadcast Text Modification")
            Exit Sub
        End If
    Next
    
    s = GetPlayerName(Index) & ": " & Msg
    Call AddLog(s, PLAYER_LOG)
    Call GlobalMsg(COLOR_BROADCAST & s)
    Call TextAdd(s)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Global Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        s = "(global) " & GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, ADMIN_LOG)
        Call GlobalMsg(COLOR_GLOBAL & s)
        Call TextAdd(s)
    End If
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Admin Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
        Call AdminMsg(COLOR_ADMIN & "(admin " & GetPlayerName(Index) & ") " & Msg)
    End If
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Player Msg Text Modification")
            Exit Sub
        End If
    Next
    
    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, COLOR_TELL & GetPlayerName(Index) & " tells you, '" & Msg & "'")
            Call PlayerMsg(Index, COLOR_TELL & "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'")
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(GetPlayerName(Index), COLOR_BRIGHTRED & "Cannot message yourself.")
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    
    If TempPlayer(Index).GettingRoom = YES Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    
    ' Prevent hacking
    If Dir < DIR_NORTH Or Dir > DIR_EAST Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).CastedSpell = YES Then
        ' Check if they have already casted a spell, and if so we can't let them move
        If GetTickCount > TempPlayer(Index).AttackTimer + 1000 Then
            TempPlayer(Index).CastedSpell = NO
        Else
            Exit Sub
        End If
    End If
    
    Call PlayerMove(Index, Dir)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    
    If TempPlayer(Index).GettingRoom = YES Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Prevent hacking
    If Dir < DIR_NORTH Or Dir > DIR_EAST Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 10
    Buffer.WriteInteger SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    
    Call SendDataToRoomBut(Index, GetPlayerRoom(Index), Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Private Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim ItemNum As Long
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    If (ItemNum > 0) And (ItemNum <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
        
        ' Find out what kind of item it is
        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
        Case ITEM_TYPE_ARMOR
            If InvNum <> GetPlayerEquipmentSlot(Index, Armor) Then
                If GetPlayerStat(Index, Stats.Defense) < n Then
                    Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your defense is to low to wear this armor!  Required DEF (" & n & ")")
                    Exit Sub
                End If
                Call SetPlayerEquipmentSlot(Index, InvNum, Armor)
            Else
                Call SetPlayerEquipmentSlot(Index, 0, Armor)
            End If
            Call SendWornEquipment(Index)
            
        Case ITEM_TYPE_WEAPON
            If InvNum <> GetPlayerEquipmentSlot(Index, Weapon) Then
                If GetPlayerStat(Index, Stats.Strength) < n Then
                    Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your strength is to low to hold this weapon!  Required STR (" & n & ")")
                    Exit Sub
                End If
                Call SetPlayerEquipmentSlot(Index, InvNum, Weapon)
            Else
                Call SetPlayerEquipmentSlot(Index, 0, Weapon)
            End If
            Call SendWornEquipment(Index)
            
        Case ITEM_TYPE_HELMET
            If InvNum <> GetPlayerEquipmentSlot(Index, Helmet) Then
                If GetPlayerStat(Index, Stats.Speed) < n Then
                    Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n & ")")
                    Exit Sub
                End If
                Call SetPlayerEquipmentSlot(Index, InvNum, Helmet)
            Else
                Call SetPlayerEquipmentSlot(Index, 0, Helmet)
            End If
            Call SendWornEquipment(Index)
            
        Case ITEM_TYPE_SHIELD
            If InvNum <> GetPlayerEquipmentSlot(Index, Shield) Then
                Call SetPlayerEquipmentSlot(Index, InvNum, Shield)
            Else
                Call SetPlayerEquipmentSlot(Index, 0, Shield)
            End If
            Call SendWornEquipment(Index)
            
        Case ITEM_TYPE_POTIONADDHP
            Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.HP)
            
        Case ITEM_TYPE_POTIONADDMP
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.MP)
            
        Case ITEM_TYPE_POTIONADDSP
            Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.SP)
            
        Case ITEM_TYPE_POTIONSUBHP
            Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) - Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.HP)
            
        Case ITEM_TYPE_POTIONSUBMP
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.MP)
            
        Case ITEM_TYPE_POTIONSUBSP
            Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) - Item(ItemNum).Data1)
            Call TakeItem(Index, ItemNum, 0)
            Call SendVital(Index, Vitals.SP)
            
        Case ITEM_TYPE_SPELL
            ' Get the spell num
            n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
            
            If n > 0 Then
                ' Make sure they are the right class
                If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                    ' Make sure they are the right level
                    i = Spell(n).LevelReq
                    If i <= GetPlayerLevel(Index) Then
                        i = FindOpenSpellSlot(Index)
                        
                        ' Make sure they have an open spell slot
                        If i > 0 Then
                            ' Make sure they dont already have the spell
                            If Not HasSpell(Index, n) Then
                                Call SetPlayerSpell(Index, i, n)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, COLOR_YELLOW & "You study the spell carefully...")
                                Call PlayerMsg(Index, COLOR_WHITE & "You have learned a new spell!")
                                Call SendSpells(Index)
                            Else
                                Call PlayerMsg(Index, COLOR_BRIGHTRED & "You have already learned this spell!")
                            End If
                        Else
                            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You have learned all that you can learn!")
                        End If
                    Else
                        Call PlayerMsg(Index, COLOR_WHITE & "You must be level " & i & " to learn this spell.")
                    End If
                Else
                    Call PlayerMsg(Index, COLOR_WHITE & "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq - 1) & ".")
                End If
            Else
                Call PlayerMsg(Index, COLOR_WHITE & "This scroll is not connected to a spell, please inform an admin!")
            End If
            
        End Select
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim tInd As Integer
    
    If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER Then
    
        tInd = TempPlayer(Index).Target
        ' Try to attack a player
        ' Make sure we dont try to attack ourselves
        If tInd <> Index Then
            ' Can we attack the player?
            If CanAttackPlayer(Index, tInd) Then
                If Not CanPlayerBlockHit(tInd) Then
                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(tInd)
                    Else
                        n = GetPlayerDamage(Index)
                        Damage = n + Int(Rnd * (n \ 2)) + 1 - GetPlayerProtection(tInd)
                        Call PlayerMsg(Index, COLOR_BRIGHTCYAN & "You feel a surge of energy upon swinging!")
                        Call PlayerMsg(tInd, COLOR_BRIGHTCYAN & GetPlayerName(Index) & " swings with enormous might!")
                    End If
                    
                    Call AttackPlayer(Index, tInd, Damage)
                    
                Else
                    Call PlayerMsg(Index, COLOR_BRIGHTCYAN & GetPlayerName(TempIndex) & "'s " & Trim$(Item(GetPlayerInvItemNum(tInd, GetPlayerEquipmentSlot(tInd, Shield))).Name) & " has blocked your hit!")
                    Call PlayerMsg(tInd, COLOR_BRIGHTCYAN & "Your " & Trim$(Item(GetPlayerInvItemNum(tInd, GetPlayerEquipmentSlot(tInd, Shield))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!")
                End If
                
                Exit Sub
            End If
        End If
    
    ElseIf TempPlayer(Index).TargetType = TARGET_TYPE_NPC Then
        
        tInd = TempPlayer(Index).Target
        ' Can we attack the Npc?
        If CanAttackNpc(Index, tInd) Then
            ' Get the damage we can do
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - (Npc(RoomNpc(GetPlayerRoom(Index), tInd).Num).Stat(Stats.Defense) \ 2)
            Else
                n = GetPlayerDamage(Index)
                Damage = n + Int(Rnd * (n \ 2)) + 1 - (Npc(RoomNpc(GetPlayerRoom(Index), tInd).Num).Stat(Stats.Defense) \ 2)
                Call PlayerMsg(Index, COLOR_BRIGHTCYAN & "You feel a surge of energy upon swinging!")
            End If
            
            If Damage > 0 Then
                Call AttackNpc(Index, tInd, Damage)
            Else
                Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your attack does nothing.")
            End If
            Exit Sub
        End If
        
    End If
    
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Private Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PointType As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    PointType = Buffer.ReadLong
    
    ' Prevent hacking
    If (PointType < 0) Or (PointType > 3) Then
        Call HackingAttempt(Index, "Invalid Point Type")
        Exit Sub
    End If
    
    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
        
        ' Everything is ok
        Select Case PointType
        Case 0
            Call SetPlayerStat(Index, Stats.Strength, GetPlayerStat(Index, Stats.Strength) + 1)
            Call PlayerMsg(Index, COLOR_WHITE & "You have gained more strength!")
        Case 1
            Call SetPlayerStat(Index, Stats.Defense, GetPlayerStat(Index, Stats.Defense) + 1)
            Call PlayerMsg(Index, COLOR_WHITE & "You have gained more defense!")
        Case 2
            Call SetPlayerStat(Index, Stats.Magic, GetPlayerStat(Index, Stats.Magic) + 1)
            Call PlayerMsg(Index, COLOR_WHITE & "You have gained more magic abilities!")
        Case 3
            Call SetPlayerStat(Index, Stats.Speed, GetPlayerStat(Index, Stats.Speed) + 1)
            Call PlayerMsg(Index, COLOR_WHITE & "You have gained more speed!")
        End Select
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "You have no skill points to train with!")
    End If
    
    ' Send the update
    Call SendStats(Index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Private Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    i = FindPlayer(Name)
    If i > 0 Then
        Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i))
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "-=- Stats for " & GetPlayerName(i) & " -=-")
            Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i))
            Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "HP: " & GetPlayerVital(i, Vitals.HP) & "/" & GetPlayerMaxVital(i, Vitals.HP) & "  MP: " & GetPlayerVital(i, Vitals.MP) & "/" & GetPlayerMaxVital(i, Vitals.MP) & "  SP: " & GetPlayerVital(i, Vitals.SP) & "/" & GetPlayerMaxVital(i, Vitals.SP) & "  Stamina: " & GetPlayerVital(i, Vitals.Stamina) & "/" & GetPlayerMaxVital(i, Vitals.Stamina))
            Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "Strength: " & GetPlayerStat(i, Stats.Strength) & "  Defense: " & GetPlayerStat(i, Stats.Defense) & "  Magic: " & GetPlayerStat(i, Stats.Magic) & "  Speed: " & GetPlayerStat(i, Stats.Speed))
            n = (GetPlayerStat(i, Stats.Strength) \ 2) + (GetPlayerLevel(i) \ 2)
            i = (GetPlayerStat(i, Stats.Defense) \ 2) + (GetPlayerLevel(i) \ 2)
            If n > 100 Then n = 100
            If i > 100 Then i = 100
            Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
    End If
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The player
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerRoom(n))
            Call PlayerMsg(n, COLOR_BRIGHTBLUE & GetPlayerName(Index) & " has warped to you.")
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "You have been warped to " & GetPlayerName(n) & ".")
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", Room #" & GetPlayerRoom(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "You cannot warp to yourself.")
    End If
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The player
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerRoom(Index))
            Call PlayerMsg(n, COLOR_BRIGHTBLUE & "You have been summoned by " & GetPlayerName(Index) & ".")
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & GetPlayerName(n) & " has been summoned.")
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, Room #" & GetPlayerRoom(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "You cannot warp yourself to yourself!")
    End If
End Sub

' ::::::::::::::::::::::::
' :: Warp to Room packet ::
' ::::::::::::::::::::::::
Private Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The Room
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_ROOMS Then
        Call HackingAttempt(Index, "Invalid Room")
        Exit Sub
    End If
    
    Call PlayerWarp(Index, n)
    Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "You have been warped to Room #" & n)
    Call AddLog(GetPlayerName(Index) & " warped to Room #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set Avatar packet ::
' :::::::::::::::::::::::
Private Sub HandleSetAvatar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The Avatar
    n = Buffer.ReadLong
    
    Call SetPlayerAvatar(Index, n)
    Call SendPlayerData(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    
    Call PlayerMsg(Index, COLOR_WHITE & "-=- Stats for " & GetPlayerName(Index) & " -=-")
    Call PlayerMsg(Index, COLOR_WHITE & "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index))
    Call PlayerMsg(Index, COLOR_WHITE & "HP: " & GetPlayerVital(Index, Vitals.HP) & "/" & GetPlayerMaxVital(Index, Vitals.HP) & "  MP: " & GetPlayerVital(Index, Vitals.MP) & "/" & GetPlayerMaxVital(Index, Vitals.MP) & "  SP: " & GetPlayerVital(Index, Vitals.SP) & "/" & GetPlayerMaxVital(Index, Vitals.SP) & "  Stamina: " & GetPlayerVital(Index, Vitals.Stamina) & "/" & GetPlayerMaxVital(Index, Vitals.Stamina))
    Call PlayerMsg(Index, COLOR_WHITE & "STR: " & GetPlayerStat(Index, Stats.Strength) & "  DEF: " & GetPlayerStat(Index, Stats.Defense) & "  MAGI: " & GetPlayerStat(Index, Stats.Magic) & "  Speed: " & GetPlayerStat(Index, Stats.Speed))
    n = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
    i = (GetPlayerStat(Index, Stats.Defense) \ 2) + (GetPlayerLevel(Index) \ 2)
    If n > 100 Then n = 100
    If i > 100 Then i = 100
    Call PlayerMsg(Index, COLOR_WHITE & "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%")
    
    Call SendStats(Index)
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new Room ::
' ::::::::::::::::::::::::::::::::::
Private Sub HandleRequestNewRoom(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    
    ' Prevent hacking
    If Dir < DIR_NORTH Or Dir > DIR_EAST Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Call PlayerMove(Index, Dir)
End Sub

' :::::::::::::::::::::
' :: Room data packet ::
' :::::::::::::::::::::
Private Sub HandleRoomData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim RoomNum As Long
    Dim RoomSize As Long
    Dim RoomData() As Byte
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    RoomNum = GetPlayerRoom(Index)
    
    i = Room(RoomNum).Revision + 1
    
    Call ClearRoom(RoomNum)
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    RoomSize = LenB(Room(RoomNum))
    ReDim RoomData(RoomSize - 1)
    RoomData = Buffer.ReadBytes(RoomSize)
    CopyMemory ByVal VarPtr(Room(RoomNum)), ByVal VarPtr(RoomData(0)), RoomSize
    
    ' set the new revision
    Room(RoomNum).Revision = i
    
    For i = 1 To MAX_ROOM_NPCS
        Call ClearRoomNpc(i, RoomNum)
    Next
    
    Call SendRoomNpcsToRoom(RoomNum)
    Call SpawnRoomNpcs(RoomNum)
    
    ' Clear out it all
    For i = 1 To MAX_ROOM_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerRoom(Index))
        Call ClearRoomItem(i, RoomNum)
    Next
    
    ' Respawn
    Call SpawnRoomItems(GetPlayerRoom(Index))
    
    ' Save the Room
    Call SaveRoom(RoomNum)
    
    Call RoomCache_Create(RoomNum)
    
    ' Refresh Room for everyone online
    For i = 1 To TotalPlayersOnline
        i = PlayersOnline(i)
        If IsPlaying(i) Then
            Call SendRoom(i, RoomNum)
        End If
    Next
End Sub

' ::::::::::::::::::::::::::::
' :: Need Room yes/no packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleNeedRoom(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As Byte
    Dim i As Byte
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Get yes/no value
    s = Buffer.ReadByte
    
    Set Buffer = Nothing
    
    ' Check if Room data is needed to be sent
    If s = 1 Then
        Call SendRoom(Index, GetPlayerRoom(Index))
    End If
    
    
    For i = 1 To MAX_ROOMS
        Call SendRoomItemsTo(Index, i)
        Call SendRoomNpcsTo(Index, i)
    Next i
    Call SendJoinRoom(Index)
    TempPlayer(Index).GettingRoom = NO
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SRoomDone
    Call SendDataTo(Index, Buffer.ToArray())
    
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleRoomGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemSel As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ItemSel = Buffer.ReadLong
    Call PlayerRoomGetItem(Index, ItemSel)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleRoomDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Amount > GetPlayerInvItemValue(Index, InvNum) Then
        Call HackingAttempt(Index, "Item amount modification")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
        If Amount <= 0 Then
            'Call HackingAttempt(Index, "Trying to drop 0 amount of currency")
            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0) ' remove item
            Exit Sub
        End If
    End If
    
    Call PlayerRoomDropItem(Index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn Room packet ::
' ::::::::::::::::::::::::
Private Sub HandleRoomRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' Clear out it all
    For i = 1 To MAX_ROOM_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerRoom(Index))
        Call ClearRoomItem(i, GetPlayerRoom(Index))
    Next
    
    ' Respawn
    Call SpawnRoomItems(GetPlayerRoom(Index))
    
    ' Respawn NpcS
    For i = 1 To MAX_ROOM_NPCS
        Call SpawnNpc(i, GetPlayerRoom(Index))
    Next
    
    Call PlayerMsg(Index, COLOR_BLUE & "Room respawned.")
    Call AddLog(GetPlayerName(Index) & " has respawned Room #" & GetPlayerRoom(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Room report packet ::
' :::::::::::::::::::::::
Private Sub HandleRoomReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tRoomStart As Long
    Dim tRoomEnd As Long
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    s = "Free Rooms: "
    tRoomStart = 1
    tRoomEnd = 1
    
    For i = 1 To MAX_ROOMS
        If LenB(Trim$(Room(i).Name)) = 0 Then
            tRoomEnd = tRoomEnd + 1
        Else
            If tRoomEnd - tRoomStart > 0 Then
                s = s & Trim$(CStr(tRoomStart)) & "-" & Trim$(CStr(tRoomEnd - 1)) & ", "
            End If
            tRoomStart = i + 1
            tRoomEnd = i + 1
        End If
    Next
    
    s = s & Trim$(CStr(tRoomStart)) & "-" & Trim$(CStr(tRoomEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    
    Call PlayerMsg(Index, COLOR_BROWN & s)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Private Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The player index
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
            
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(COLOR_WHITE & GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!")
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, COLOR_WHITE & "That is a higher or same access admin then you!")
            End If
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "You cannot kick yourself!")
    End If
End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Private Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F
    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        
        Call PlayerMsg(Index, COLOR_WHITE & n & ": Banned IP " & s & " by " & Name)
        n = n + 1
    Loop
    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Private Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim FileName As String
    Dim File As Long
    Dim F As Long
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    FileName = App.Path & "\data\banlist.txt"
    
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    Kill FileName
    
    Call PlayerMsg(Index, COLOR_WHITE & "Ban list destroyed.")
    
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The player index
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, COLOR_WHITE & "That is a higher or same access admin then you!")
            End If
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "You cannot ban yourself!")
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Room packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditRoom(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SEditRoom
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SItemEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The item #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
    Call SendEditItemTo(Index, n)
End Sub

Private Sub HandleDelete(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Editor As Byte
    Dim EditorIndex As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Editor = Buffer.ReadByte
    EditorIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    Select Case Editor
        
    Case EDITOR_ITEM
        ' Prevent hacking
        If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call ClearItem(EditorIndex)
        
        Call SendUpdateItemToAll(EditorIndex)
        Call SaveItem(EditorIndex)
        Call AddLog(GetPlayerName(Index) & "Deleted item #" & EditorIndex & ".", ADMIN_LOG)
        
    Case EDITOR_Npc
        ' Prevent hacking
        If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid Npc Index")
            Exit Sub
        End If
        
        Call ClearNpc(EditorIndex)
        
        Call SendUpdateNpcToAll(EditorIndex)
        Call SaveNpc(EditorIndex)
        Call AddLog(GetPlayerName(Index) & "Deleted Npc #" & EditorIndex & ".", ADMIN_LOG)
        
    Case EDITOR_SPELL
        ' Prevent hacking
        If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call ClearSpell(EditorIndex)
        
        Call SendUpdateSpellToAll(EditorIndex)
        Call SaveSpell(EditorIndex)
        Call AddLog(GetPlayerName(Index) & "Deleted spell #" & EditorIndex & ".", ADMIN_LOG)
        
    Case EDITOR_SHOP
        ' Prevent hacking
        If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call ClearShop(EditorIndex)
        
        Call SendUpdateShopToAll(EditorIndex)
        Call SaveShop(EditorIndex)
        Call AddLog(GetPlayerName(Index) & "Deleted shop #" & EditorIndex & ".", ADMIN_LOG)
    End Select
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SREditor
    Call SendDataTo(Index, Buffer.ToArray())
    
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ItemNum As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
    
    ' Save it
    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)
    Call AddLog(GetPlayerName(Index) & " saved item #" & ItemNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Npc packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SNpcEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Edit Npc packet ::
' :::::::::::::::::::::
Private Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The Npc #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid Npc Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing Npc #" & n & ".", ADMIN_LOG)
    Call SendEditNpcTo(Index, n)
End Sub

' :::::::::::::::::::::
' :: Save Npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    
    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid Npc Index")
        Exit Sub
    End If
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
    
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SShopEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The shop #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
    Call SendEditShopTo(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ShopNum As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRequestEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SSpellEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
' :::::::::::::::::::::::
Private Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The spell #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
    Call SendEditSpellTo(Index, n)
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Private Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Spell #
    SpellNum = Buffer.ReadLong
    
    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(Index) & " saving spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Private Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Trying to use powers not available")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' The index
    n = FindPlayer(Buffer.ReadString)
    ' The access
    i = Buffer.ReadLong
    
    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then
        ' Check if player is on
        If n > 0 Then
            
            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, COLOR_RED & "Invalid access level.")
                Exit Sub
            End If
            
            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(COLOR_BRIGHTBLUE & GetPlayerName(n) & " has been blessed with administrative access.")
            End If
            
            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
        End If
    Else
        Call PlayerMsg(Index, COLOR_RED & "Invalid access level.")
    End If
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Private Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Private Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MOTD = Buffer.ReadString
    Call PutVar(App.Path & "\data\motd.ini", "MOTD", "Msg", MOTD)
    Call GlobalMsg(COLOR_BRIGHTCYAN & "MOTD changed to: " & MOTD)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

' ::::::::::::::::::
' :: Trade packet ::
' ::::::::::::::::::
Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Room(GetPlayerRoom(Index)).Shop > 0 Then
        Call SendTrade(Index, Room(GetPlayerRoom(Index)).Shop)
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "There is no shop here.")
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Trade request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim X As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Trade num
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If (n <= 0) Or (n > MAX_TRADES) Then
        Call HackingAttempt(Index, "Trade Request Modification")
        Exit Sub
    End If
    
    ' Index for shop
    i = Room(GetPlayerRoom(Index)).Shop
    
    ' Check if inv full
    X = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
    If X = 0 Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Trade unsuccessful, inventory full.")
        Exit Sub
    End If
    
    ' Check if they have the item
    If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
        Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
        Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
        Call PlayerMsg(Index, COLOR_YELLOW & "The trade was successful!")
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Trade unsuccessful.")
    End If
End Sub

' :::::::::::::::::::::
' :: Fix item packet ::
' :::::::::::::::::::::
Private Sub HandleFixItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim ItemNum As Long
    Dim DurNeeded As Long
    Dim GoldNeeded As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Inv num
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n <= 0 Or n > MAX_INV Then
        Call HackingAttempt(Index, "Fix item modification")
        Exit Sub
    End If
    
    ' check for bad data
    If GetPlayerInvItemNum(Index, n) <= 0 Or GetPlayerInvItemNum(Index, n) > MAX_ITEMS Then
        Exit Sub
    End If
    
    ' Make sure its a equipable item
    If Item(GetPlayerInvItemNum(Index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, n)).Type > ITEM_TYPE_SHIELD Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "You can only fix weapons, armors, helmets, and shields.")
        Exit Sub
    End If
    
    ' Check if they have a full inventory
    If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, n)) <= 0 Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "You have no inventory space left!")
        Exit Sub
    End If
    
    ' Now check the rate of pay
    ItemNum = GetPlayerInvItemNum(Index, n)
    i = (Item(GetPlayerInvItemNum(Index, n)).Data2 \ 5)
    If i <= 0 Then i = 1
    
    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
    GoldNeeded = Int(DurNeeded * i / 2)
    If GoldNeeded <= 0 Then GoldNeeded = 1
    
    ' Check if they even need it repaired
    If DurNeeded <= 0 Then
        Call PlayerMsg(Index, COLOR_WHITE & "This item is in perfect condition!")
        Exit Sub
    End If
    
    ' Check if they have enough for at least one point
    If HasItem(Index, 1) >= i Then
        ' Check if they have enough for a total restoration
        If HasItem(Index, 1) >= GoldNeeded Then
            Call TakeItem(Index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "Item has been totally restored for " & GoldNeeded & " gold!")
        Else
            ' They dont so restore as much as we can
            DurNeeded = (HasItem(Index, 1) / i)
            GoldNeeded = Int(DurNeeded * i \ 2)
            If GoldNeeded <= 0 Then GoldNeeded = 1
            
            Call TakeItem(Index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "Item has been partially fixed for " & GoldNeeded & " gold!")
        End If
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Insufficient gold to fix this item!")
    End If
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Private Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim pName As String, mNPCNum As Long, mItemNum As Long, bType As Byte
    Dim X As Long, y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    bType = Buffer.ReadByte
    
    Select Case bType
        Case 0 ' Players
            pName = Buffer.ReadString
            
            ' Check for a player
            For i = 1 To TotalPlayersOnline
                If GetPlayerName(PlayersOnline(i)) = pName Then
                    If GetPlayerRoom(Index) = GetPlayerRoom(PlayersOnline(i)) Then
                                
                        ' Consider the player
                        If PlayersOnline(i) <> Index Then
                            
                            If GetPlayerLevel(PlayersOnline(i)) >= GetPlayerLevel(Index) + 5 Then
                                Call PlayerMsg(Index, COLOR_BRIGHTRED & "You wouldn't stand a chance.")
                            Else
                                If GetPlayerLevel(PlayersOnline(i)) > GetPlayerLevel(Index) Then
                                    Call PlayerMsg(Index, COLOR_YELLOW & "This one seems to have an advantage over you.")
                                Else
                                    If GetPlayerLevel(PlayersOnline(i)) = GetPlayerLevel(Index) Then
                                        Call PlayerMsg(Index, COLOR_WHITE & "This would be an even fight.")
                                    Else
                                        If GetPlayerLevel(Index) >= GetPlayerLevel(PlayersOnline(i)) + 5 Then
                                            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "You could slaughter that player.")
                                        Else
                                            If GetPlayerLevel(Index) > GetPlayerLevel(PlayersOnline(i)) Then
                                                Call PlayerMsg(Index, COLOR_YELLOW & "You would have an advantage over that player.")
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        
                        ' Change target
                        TempPlayer(Index).Target = PlayersOnline(i)
                        TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
                        'Call PlayerMsg(Index, COLOR_YELLOW & "Your target is now " & GetPlayerName(PlayersOnline(i)) & ".")
                        Exit Sub
                    End If
                End If
            Next
            
        Case 1 ' NPCs
            mNPCNum = Buffer.ReadLong
            ' Check for an Npc
            If RoomNpc(GetPlayerRoom(Index), mNPCNum).Num > 0 Then
                ' Change target
                TempPlayer(Index).Target = mNPCNum
                TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                'Call PlayerMsg(Index, COLOR_YELLOW & "Your target is now a " & Trim$(Npc(RoomNpc(GetPlayerRoom(Index), mNPCNum).Num).Name) & ".")
                Exit Sub
            End If
            
        Case 2 ' Items
            mItemNum = Buffer.ReadLong
            ' Check for an item
            If RoomItem(GetPlayerRoom(Index), mItemNum).Num > 0 Then
                'Call PlayerMsg(Index, COLOR_YELLOW & "You see a " & Trim$(Item(RoomItem(GetPlayerRoom(Index), mItemNum).Num).Name) & ".")
                Exit Sub
            End If
    End Select
End Sub

' ::::::::::::::::::
' :: Party packet ::
' ::::::::::::::::::
Private Sub HandleParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    n = FindPlayer(Buffer.ReadString)
    
    ' Prevent partying with self
    If n = Index Then
        Exit Sub
    End If
    
    ' Check for a previous party and if so drop it
    If TempPlayer(Index).InParty = YES Then
        Call PlayerMsg(Index, COLOR_PINK & "You are already in a party!")
        Exit Sub
    End If
    
    If n > 0 Then
        ' Check if its an admin
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "You can't join a party, you are an admin!")
            Exit Sub
        End If
        
        If GetPlayerAccess(n) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "Admins cannot join parties!")
            Exit Sub
        End If
        
        ' Make sure they are in right level range
        If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
            Call PlayerMsg(Index, COLOR_PINK & "There is more then a 5 level gap between you two, party failed.")
            Exit Sub
        End If
        
        ' Check to see if player is already in a party
        If TempPlayer(n).InParty = NO Then
            Call PlayerMsg(Index, COLOR_PINK & "Party request has been sent to " & GetPlayerName(n) & ".")
            Call PlayerMsg(n, COLOR_PINK & GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.")
            
            TempPlayer(Index).PartyStarter = YES
            TempPlayer(Index).PartyPlayer = n
            TempPlayer(n).PartyPlayer = Index
        Else
            Call PlayerMsg(Index, COLOR_PINK & "Player is already in a party!")
        End If
    Else
        Call PlayerMsg(Index, COLOR_WHITE & "Player is not online.")
    End If
End Sub

' :::::::::::::::::::::::
' :: Join party packet ::
' :::::::::::::::::::::::
Private Sub HandleJoinParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    
    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        ' Check to make sure they aren't the starter
        If TempPlayer(Index).PartyStarter = NO Then
            ' Check to make sure that each of there party players match
            If TempPlayer(n).PartyPlayer = Index Then
                Call PlayerMsg(Index, COLOR_PINK & "You have joined " & GetPlayerName(n) & "'s party!")
                Call PlayerMsg(n, COLOR_PINK & GetPlayerName(Index) & " has joined your party!")
                
                TempPlayer(Index).InParty = YES
                TempPlayer(n).InParty = YES
            Else
                Call PlayerMsg(Index, COLOR_PINK & "Party failed.")
            End If
        Else
            Call PlayerMsg(Index, COLOR_PINK & "You have not been invited to join a party!")
        End If
    Else
        Call PlayerMsg(Index, COLOR_PINK & "You have not been invited into a party!")
    End If
End Sub

' ::::::::::::::::::::::::
' :: Leave party packet ::
' ::::::::::::::::::::::::
Private Sub HandleLeaveParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    
    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        If TempPlayer(Index).InParty = YES Then
            Call PlayerMsg(Index, COLOR_PINK & "You have left the party.")
            Call PlayerMsg(n, COLOR_PINK & GetPlayerName(Index) & " has left the party.")
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        Else
            Call PlayerMsg(Index, COLOR_PINK & "Declined party request.")
            Call PlayerMsg(n, COLOR_PINK & GetPlayerName(Index) & " declined your request.")
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        End If
    Else
        Call PlayerMsg(Index, COLOR_PINK & "You are not in a party!")
    End If
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Private Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Private Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Spell slot
    n = Buffer.ReadLong
    
    Call CastSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Private Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::
' :: Sync ::
' ::::::::::

Private Sub HandleSync(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteInteger SSync
    Buffer.WriteLong GetPlayerRoom(Index)
    
    Call SendDataTo(Index, Buffer.ToArray())
    
End Sub

Private Sub HandleRoomReqs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_ROOMS
        If Buffer.ReadByte = 1 Then
            SendRoom Index, i
        End If
    Next i
    
End Sub

Private Sub HandleSleepInn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long
    
    If Room(GetPlayerRoom(Index)).Moral <> ROOM_MORAL_INN Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "There is no Inn here. You cannot sleep!")
        Exit Sub
    End If
        
    For i = 1 To MAX_INV
        Select Case GetPlayerInvItemName(Index, i)
        Case "Gold"
            If GetPlayerInvItemValue(Index, i) >= Val(GetPlayerLevel(Index) * 10) Then
                Call TakeItem(Index, i, Val(GetPlayerLevel(Index) * 10))
                For n = 1 To Vitals.Vital_Count - 1
                    Call SetPlayerVital(Index, n, GetPlayerMaxVital(Index, n))
                    Call SendVital(Index, n)
                Next
                Call PlayerMsg(Index, COLOR_BRIGHTGREEN & "You sleep and wake up feeling refreshed!")
                Exit Sub
            ElseIf GetPlayerInvItemValue(Index, i) < Val(GetPlayerLevel(Index) * 10) Then
                Call PlayerMsg(Index, COLOR_BRIGHTRED & "You do not have enough money to sleep here!")
                Exit Sub
            End If
        Case Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You do not have any gold with which to pay!")
            Exit Sub
        End Select
    Next i
End Sub

Private Sub HandleCreateGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = Buffer.ReadString
    
    If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
        For i = 1 To High_Index
            If user = LCase(GetPlayerName(i)) Then
                Call SetPlayerGuild(i, Guild)
                Call SetPlayerGAccess(i, 2)
                Call SendPlayerData(i)
                
                Exit Sub
            End If
        Next
    End If
    
End Sub

Private Sub HandleRemoveFromGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = Buffer.ReadString
    
    If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, 0)
                    Call SetPlayerGuild(i, "")
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
    
End Sub

Private Sub HandleGuildInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = "" Then
                    Call SetPlayerGAccess(i, 1)
                    Call SetPlayerGuild(i, Guild)
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleGuildKick(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, 0)
                    Call SetPlayerGuild(i, "")
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleGuildPromote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    Dim Access As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Access = Buffer.ReadLong
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        If Access > GetPlayerGAccess(Index) Then
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You cannot set access higher than your own.")
            Exit Sub
        End If
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, Access)
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleLeaveGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim user As String
    Dim Guild As String
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    'user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 1 Then
        Call SetPlayerGAccess(Index, 0)
        Call SetPlayerGuild(Index, "")
        Call SendPlayerData(Index)
    End If
End Sub
