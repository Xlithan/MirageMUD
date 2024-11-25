Attribute VB_Name = "GameLogic"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public FpsLowTick  As Long

Public FpsLowCount As Long

Public Sub GameLoop()

    Dim i, n As Long
    
    Dim TickFPS   As Long

    Dim FPS       As Long
    
    Dim Tick      As Long
    
    Dim WalkTimer As Long

    Dim tmr25     As Long

    Dim tmr10000  As Long
    
    vbQuote = Chr(34)
    
    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount
        
        If tmr25 < Tick Then
            
            InGame = IsConnected
            
            If SentSync = False Then Call SyncPacket
            
            If GetForegroundWindow() = frmMainGame.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            If CanMoveNow Then
                Call CheckAttack   ' Check to see if player is trying to attack
            End If
            
            tmr25 = Tick + 25
        End If

        DoEvents
        
    Loop
    
    frmMainGame.Visible = False
    
    If isLogging Then
        frmMainGame.txtChat = vbNullString
        frmMainGame.txtMyChat = vbNullString
        isLogging = False
        frmMainMenu.Visible = False
        GettingRoom = True
    Else
        ' Shutdown the game
        frmSendGetData.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If
    
End Sub

Public Sub CheckAttack()

    Dim Buffer As clsBuffer
    
    If ControlDown Then
        If Player(MyIndex).AttackTimer + 1000 < GetTickCount Then
            Player(MyIndex).AttackTimer = GetTickCount
            
            Set Buffer = New clsBuffer
            Buffer.PreAllocate 2
            Buffer.WriteInteger CAttack
            Call SendData(Buffer.ToArray())
        End If
    End If
    
End Sub

Private Function CanMove() As Boolean

    Dim d As Long
    
    CanMove = True
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False

            Exit Function

        End If
    End If

End Function

Public Sub CheckMovement()
        
    If CanMove Then
        
        Select Case GetPlayerDir(MyIndex)

            Case DIR_NORTH
                Call SendPlayerMove

            Case DIR_SOUTH
                Call SendPlayerMove
            
            Case DIR_WEST
                Call SendPlayerMove
            
            Case DIR_EAST
                Call SendPlayerMove
            
        End Select
        
    End If
    
End Sub

Public Sub UpdateInventory()

    Dim i As Long
    
    frmMainGame.lstInv.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else

                ' Check if this item is being worn
                If GetPlayerEquipmentSlot(MyIndex, Weapon) = i Or GetPlayerEquipmentSlot(MyIndex, Armor) = i Or GetPlayerEquipmentSlot(MyIndex, Helmet) = i Or GetPlayerEquipmentSlot(MyIndex, Shield) = i Then
                    frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
                Else
                    frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                End If
            End If

        Else
            frmMainGame.lstInv.AddItem " "
        End If

    Next
    
    frmMainGame.lstInv.ListIndex = 0
End Sub

Public Sub UpdateSpells()

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CSpells
    Call SendData(Buffer.ToArray())
End Sub

Public Sub GetStats()

    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CGetStats
    Call SendData(Buffer.ToArray())
End Sub

Public Sub GetPlayersInRoom()

    Dim i As Long, n As Long
    
    PlayersInRoomHighIndex = 1
    n = 1
    
    ReDim PlayersInRoom(1 To MAX_PLAYERS)
    
    frmMainGame.lstPlayers.Clear
    
    For i = 1 To High_Index

        If IsPlaying(i) Then
            If GetPlayerRoom(i) = GetPlayerRoom(MyIndex) Then
                PlayersInRoom(PlayersInRoomHighIndex) = i
                PlayersInRoomHighIndex = PlayersInRoomHighIndex + 1
                frmMainGame.lstPlayers.AddItem GetPlayerName(i)
                PlayerLst(n) = i
                n = n + 1
            End If
        End If

    Next
    
    PlayersInRoomHighIndex = PlayersInRoomHighIndex - 1
End Sub

Public Sub PlayerSearch(ByVal pName As String)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 10
    Buffer.WriteInteger CSearch
    Buffer.WriteByte 0
    Buffer.WriteString pName
    Call SendData(Buffer.ToArray())

End Sub

Public Sub NPCSearch(ByVal mNPCNum As String)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 10
    Buffer.WriteInteger CSearch
    Buffer.WriteByte 1
    Buffer.WriteLong mNPCNum
    Call SendData(Buffer.ToArray())

End Sub

Public Sub ItemSearch(ByVal mItemNum As String)

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 10
    Buffer.WriteInteger CSearch
    Buffer.WriteByte 2
    Buffer.WriteLong mItemNum
    Call SendData(Buffer.ToArray())

End Sub

Public Sub UseItem()

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then

        Exit Sub

    End If
    
    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub CastSpell()

    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If SpellSelected < 1 Or SpellSelected > MAX_SPELLS Then

        Exit Sub

    End If
    
    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(SpellSelected).MPReq Then
        Call AddText(COLOR_BRIGHTRED & "Not enough MP to cast " & Trim$(Spell(SpellSelected).name) & ".")

        Exit Sub

    End If
    
    If PlayerSpells(SpellSelected) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Set Buffer = New clsBuffer
            Buffer.PreAllocate 6
            Buffer.WriteInteger CCast
            Buffer.WriteLong SpellSelected
            Call SendData(Buffer.ToArray())
            Player(MyIndex).Attacking = 1
            Player(MyIndex).AttackTimer = GetTickCount
            Player(MyIndex).CastedSpell = YES
        End If

    Else
        Call AddText(COLOR_BRIGHTRED & "No spell here.")
    End If

End Sub

Public Sub InitRoomData()

    Dim i         As Long

    Dim MusicFile As String
    
    MusicFile = Trim$(CStr(Room.Music))
    
    '    ' get high NPC index
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
End Sub

Public Sub DevMsg(ByVal Text As String)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text)
        End If
    End If

    Debug.Print Text
End Sub
