Attribute VB_Name = "Input"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' keyboard input declares
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetForegroundWindow Lib "user32" () As Long

' player input text buffer
Public MyText            As String

' Game direction vars
Public DirNorth             As Boolean

Public DirSouth           As Boolean

Public DirWest           As Boolean

Public DirEast          As Boolean

Public ShiftDown         As Boolean

Public ControlDown       As Boolean

' Key constants
Private Const VK_UP      As Long = &H26

Private Const VK_DOWN    As Long = &H28

Private Const VK_LEFT    As Long = &H25

Private Const VK_RIGHT   As Long = &H27

Private Const VK_SHIFT   As Long = &H10

'Private Const VK_RETURN As Long = &HD ' not used
Private Const VK_CONTROL As Long = &H11

Public Sub CheckInputKeys()
    
    ' Check to make sure they aren't trying to auto do anything
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    If frmMainGame.WindowState = vbMinimized Then Exit Sub
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
End Sub

Public Sub CheckRoomGetItem()

    Dim Buffer As clsBuffer

    If GetTickCount > Player(MyIndex).RoomGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Set Buffer = New clsBuffer
            Buffer.PreAllocate 2
            Buffer.WriteInteger CRoomGetItem
            Buffer.WriteLong ItemSel
            Player(MyIndex).RoomGetTimer = GetTickCount
            Call SendData(Buffer.ToArray())
        End If
    End If

End Sub

' Processes input from player
Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)

    Dim ChatText  As String

    Dim name      As String

    Dim i         As Long

    Dim n         As Long

    Dim Command() As String

    Dim Buffer    As clsBuffer
    
    ' Remove any instances of BB Code from chat
    If Player(MyIndex).Access < 4 Then
        Call RemoveBBCode(MyText)
    End If
    
    ChatText = Trim$(MyText)
    
    If LenB(ChatText) = 0 Then Exit Sub
    
    MyText = LCase$(ChatText)
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        ' Movement
        'Move North
        If ChatText = "n" Or ChatText = "north" Then

            If Len(ChatText) > 0 Then
                DirNorth = True
                DirSouth = False
                DirWest = False
                DirEast = False
                Call SetPlayerDir(MyIndex, DIR_NORTH)
            End If
            
            If CanMoveNow Then
                Call CheckMovement
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        'Move East
        If ChatText = "e" Or ChatText = "east" Then

            If Len(ChatText) > 0 Then
                DirNorth = False
                DirSouth = False
                DirWest = False
                DirEast = True
                Call SetPlayerDir(MyIndex, DIR_EAST)
            End If
            
            If CanMoveNow Then
                Call CheckMovement
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        'Move South
        If ChatText = "s" Or ChatText = "south" Then

            If Len(ChatText) > 0 Then
                DirNorth = False
                DirSouth = True
                DirWest = False
                DirEast = False
                Call SetPlayerDir(MyIndex, DIR_SOUTH)
            End If
            
            If CanMoveNow Then
                Call CheckMovement
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        'Move West
        If ChatText = "w" Or ChatText = "west" Then

            If Len(ChatText) > 0 Then
                DirNorth = False
                DirSouth = False
                DirWest = True
                DirEast = False
                Call SetPlayerDir(MyIndex, DIR_WEST)
            End If
            
            If CanMoveNow Then
                Call CheckMovement
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        ' Broadcast message
        If Left$(ChatText, 1) = "'" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call BroadcastMsg(ChatText)
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call EmoteMsg(ChatText)
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        ' Player message
        If Left$(ChatText, 1) = "!" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            name = vbNullString
            
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)

                If Mid$(ChatText, i, 1) <> Space(1) Then
                    name = name & Mid$(ChatText, i, 1)
                Else

                    Exit For

                End If

            Next
            
            ChatText = Mid$(ChatText, i, Len(ChatText) - 1)
            
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText(COLOR_ALERT & "Usage: !playername (message)")
            End If

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        ' Global Message
        If Left$(ChatText, 1) = vbQuote Then
            If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

                If Len(ChatText) > 0 Then
                    Call GlobalMsg(ChatText)
                End If

                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString

                Exit Sub

            End If
        End If
        
        ' Admin Message
        If Left$(ChatText, 1) = "=" Then
            If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

                If Len(ChatText) > 0 Then
                    Call AdminMsg(ChatText)
                End If

                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString

                Exit Sub

            End If
        End If
        
        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))
            
            Select Case Command(0)

                Case "/help"
                    Call AddText(COLOR_HELP & "Social Commands:")
                    Call AddText(COLOR_HELP & "'msghere = Broadcast Message")
                    Call AddText(COLOR_HELP & "-msghere = Emote Message")
                    Call AddText(COLOR_HELP & "!namehere msghere = Player Message")
                    Call AddText(COLOR_HELP & "Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave, /look")

                Case "/look"
                    Call AddText(COLOR_YELLOW & "<< " & COLOR_BRIGHTCYAN & Trim$(Room.name) & COLOR_YELLOW & " >>")
                    Call AddText(COLOR_BRIGHTBLUE & Trim$(Room.lDesc))
                    Call AddText(COLOR_BRIGHTBLUE & Trim$(Room.eDesc))
                
                Case "/sleep"
                    Set Buffer = New clsBuffer
                    Buffer.WriteInteger CSleepInn
                    Call SendData(Buffer.ToArray())
                    
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /info (name)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /info (name)")
                        GoTo Continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate Len(Command(1)) + 4
                    Buffer.WriteInteger CPlayerInfoRequest
                    Buffer.WriteString Command(1)
                    Call SendData(Buffer.ToArray())
                
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                
                Case "/guildinvite"

                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /guildinvite (name)")
                        GoTo Continue
                    End If
                
                    SendGuildInvite Command(1)
                
                Case "/guildkick"

                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /guildkick (name)")
                        GoTo Continue
                    End If
                
                    SendGuildKick Command(1)
                
                Case "/guildpromote"

                    If UBound(Command) < 2 Then
                        Call AddText(COLOR_ALERT & "Usage: /guildkick (name) (access)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Val(Command(2))) = False Then
                        Call AddText(COLOR_ALERT & "Usage: /guildkick (name) (access)")
                        GoTo Continue
                    End If
                
                    SendGuildPromote Command(1), Val(Command(2))
                
                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CGetStats
                    SendData Buffer.ToArray()
                
                    ' Show training
                Case "/train"
                    frmTraining.Show vbModal
                
                    ' Request stats
                Case "/trade"
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CTrade
                    SendData Buffer.ToArray()
                
                    ' Party request
                Case "/party"

                    ' Make sure they are actually sending something
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /party (name)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /party (name)")
                        GoTo Continue
                    End If
                
                    Call SendPartyRequest(Command(1))
                
                    ' Join party
                Case "/join"
                    SendJoinParty
                
                    ' Leave party
                Case "/leave"
                    SendLeaveParty
                
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    Call AddText(COLOR_HELP & "Social Commands:")
                    Call AddText(COLOR_HELP & """msghere = Global Admin Message")
                    Call AddText(COLOR_HELP & "=msghere = Private Admin Message")
                    Call AddText(COLOR_HELP & "Available Commands: /admin, /loc, /roomeditor, /warpmeto, /warptome, /warpto, /setavatar, /roomreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell, /debug")
                
                    ' Kicking a player
                Case "/kick"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /kick (name)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /kick (name)")
                        GoTo Continue
                    End If
                
                    SendKick Command(1)
                
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    BLoc = Not BLoc
                
                    ' Room Editor
                Case "/roomeditor"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRequestEditRoom
                
                    ' Warping to a player
                Case "/warpmeto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /warpmeto (name)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /warpmeto (name)")
                        GoTo Continue
                    End If
                
                    WarpMeTo Command(1)
                
                    ' Warping a player to you
                Case "/warptome"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /warptome (name)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /warptome (name)")
                        GoTo Continue
                    End If
                
                    WarpToMe Command(1)
                
                    ' Warping to a Room
                Case "/warpto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /warpto (Room #)")
                        GoTo Continue
                    End If
                
                    If Not IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /warpto (Room #)")
                        GoTo Continue
                    End If
                
                    n = CLng(Command(1))
                
                    ' Check to make sure its a valid Room #
                    If n > 0 And n <= MAX_ROOMS Then
                        Call WarpTo(n)
                    Else
                        Call AddText(COLOR_ALERT & "Invalid Room number.")
                    End If
                
                    ' Setting Avatar
                Case "/setavatar"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /setavatar (avatar #)")
                        GoTo Continue
                    End If
                
                    If Not IsNumeric(Command(1)) Then
                        Call AddText(COLOR_ALERT & "Usage: /setavatar (avatar #)")
                        GoTo Continue
                    End If
                
                    SendSetAvatar CLng(Command(1))
                
                    ' Room report
                Case "/roomreport"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CRoomReport
                    SendData Buffer.ToArray()
                
                    ' Respawn request
                Case "/respawn"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRoomRespawn
                
                    ' MOTD change
                Case "/motd"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /motd (new motd)")
                        GoTo Continue
                    End If
                
                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                
                    ' Check the ban list
                Case "/banlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendBanList
                
                    ' Banning a player
                Case "/ban"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 1 Then
                        Call AddText(COLOR_ALERT & "Usage: /ban (name)")
                        GoTo Continue
                    End If
                
                    SendBan Command(1)
                
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRequestEditItem
                
                    ' Editing npc request
                Case "/editnpc"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRequestEditNpc
                
                    ' Editing shop request
                Case "/editshop"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRequestEditShop
                
                    ' Editing spell request
                Case "/editspell"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendRequestEditSpell
                
                    ' // Creator Admin Commands //
                Case "/createguild"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 2 Then
                        Call AddText(COLOR_ALERT & "Usage: /createguild (user) (guild)")
                        GoTo Continue
                    End If
                
                    SendCreateGuild Command(1), Command(2)
                
                Case "/removefromguild"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 2 Then
                        Call AddText(COLOR_ALERT & "Usage: /removefromguild (user) (guild)")
                        GoTo Continue
                    End If
                
                    SendRemoveFromGuild Command(1), Command(2)
                
                    ' Giving another player access
                Case "/setaccess"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    If UBound(Command) < 2 Then
                        Call AddText(COLOR_ALERT & "Usage: /setaccess (name) (access)")
                        GoTo Continue
                    End If
                
                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        Call AddText(COLOR_ALERT & "Usage: /setaccess (name) (access)")
                        GoTo Continue
                    End If
                
                    SendSetAccess Command(1), CLng(Command(2))
                
                    ' Ban destroy
                Case "/destroybanlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    SendBanDestroy
                
                    ' Packet debug mode
                Case "/debug"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        Call AddText(COLOR_ALERT & "You need to be a high enough staff member to do this!")
                        GoTo Continue
                    End If
                
                    DebugMode = (Not DebugMode)
                
                Case Else
                    Call AddText(COLOR_HELP & "Not a valid command!")
                
            End Select
            
            'continue label where we go instead of exiting the sub
Continue:
            
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString

            Exit Sub

        End If
        
        ' And if neither, then add the character to the user's text buffer
        'If (KeyAscii <> vbKeyReturn) Then
        '    If (KeyAscii <> vbKeyBack) Then
        '
        '        ' Make sure the character is on standard English keyboard
        '        If KeyAscii >= 32 Then ' Asc(" ")
        '            If KeyAscii <= 126 Then ' Asc("~")
        '                MyText = MyText & ChrW$(KeyAscii)
        '            End If
        '        End If
        '
        '    End If
        'End If
        
        ' Handle when the user presses the backspace key
        'If (KeyAscii = vbKeyBack) Then
        '    If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        'End If
        
        ' Say message
        If Len(ChatText) > 0 Then
            Call SayMsg(ChatText)
        End If
        
        MyText = vbNullString
        frmMainGame.txtMyChat.Text = vbNullString

        Exit Sub

    End If
    
End Sub

