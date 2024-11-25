Attribute VB_Name = "modGameLogic"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Sub AttackNpc(ByVal Attacker As Long, ByVal RoomNpcNum As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim RoomNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or RoomNpcNum <= 0 Or RoomNpcNum > MAX_ROOM_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    RoomNum = GetPlayerRoom(Attacker)
    NpcNum = RoomNpc(RoomNum, RoomNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger SAttack
    Buffer.WriteLong Attacker
    ' Send this packet so they can see the person attacking
    Call SendDataToRoomBut(Attacker, RoomNum, Buffer.ToArray())
    
    Set Buffer = Nothing
    
    ' Check for weapon
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    If Damage >= RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points, killing it.")
            Call RoomMsgBut(RoomNum, Attacker, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with their bare fists, killing it.")
        Else
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points, killing it.")
            Call RoomMsgBut(RoomNum, Attacker, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with their " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & ", killing it.")
        End If
        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).Stat(Stats.Strength)
        DEF = Npc(NpcNum).Stat(Stats.Defense)
        Exp = STR * DEF * 2
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            
            n = TempPlayer(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
            End If
        End If
        
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, RoomNum)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        RoomNpc(RoomNum, RoomNpcNum).Num = 0
        RoomNpc(RoomNum, RoomNpcNum).SpawnWait = GetTickCount
        RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) = 0
        
        Set Buffer = New clsBuffer
        
        Buffer.PreAllocate 6 + 4
        Buffer.WriteInteger SNpcDead
        Buffer.WriteLong RoomNum
        Buffer.WriteLong RoomNpcNum
        Call SendDataToRoom(RoomNum, Buffer.ToArray())
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        If TempPlayer(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(TempPlayer(Attacker).PartyPlayer)
        End If
        
        ' Check if target is Npc that died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(Attacker).Target = RoomNpcNum Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If
    Else
        ' Npc not dead, just do the damage
        RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) = RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call RoomMsgBut(RoomNum, Attacker, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with their bare fists.")
        Else
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call RoomMsgBut(RoomNum, Attacker, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit a " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " with their " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & ".")
        End If
        
        ' Check if we should send a message
        If RoomNpc(RoomNum, RoomNpcNum).Target = 0 Then
            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                Call PlayerMsg(Attacker, COLOR_SAY & "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.")
            End If
        End If
        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).Stat(Stats.Strength)
        DEF = Npc(NpcNum).Stat(Stats.Defense)
        Exp = STR * DEF * 3
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            
            n = TempPlayer(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
            End If
        End If
        
        ' Set the Npc target to the player
        RoomNpc(RoomNum, RoomNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all Room guards come after'm
        If Npc(RoomNpc(RoomNum, RoomNpcNum).Num).Behavior = Npc_BEHAVIOR_GUARD Then
            For i = 1 To MAX_ROOM_NPCS
                If RoomNpc(RoomNum, i).Num = RoomNpc(RoomNum, RoomNpcNum).Num Then
                    RoomNpc(RoomNum, i).Target = Attacker
                End If
            Next
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Take away 1 stamina
    Call SetPlayerVital(Attacker, Vitals.Stamina, GetPlayerVital(Attacker, Vitals.Stamina) - 1)
    Call SendVital(Attacker, Vitals.Stamina)
    
    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

Public Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    n = 0
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    ' Send this packet so they can see the person attacking
'    Set Buffer = New clsBuffer
'    Buffer.PreAllocate 6
'    Buffer.WriteInteger SAttack
'    Buffer.WriteLong Attacker
'    Call SendDataToRoomBut(Attacker, GetPlayerRoom(Attacker), Buffer.ToArray())
'    Set Buffer = Nothing
    
    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with your bare fists for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit you with their bare fists for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            
            For i = 1 To TotalPlayersOnline
                If GetPlayerRoom(i) = GetPlayerRoom(Attacker) And i <> Attacker And i <> Victim Then
                    Call PlayerMsg(i, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with their bare fists.")
                End If
            Next i
        Else
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit you with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            
            For i = 1 To TotalPlayersOnline
                If GetPlayerRoom(i) = GetPlayerRoom(Attacker) And i <> Attacker And i <> Victim Then
                    Call PlayerMsg(i, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with their " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & ".")
                End If
            Next i
        End If
        
        ' Player is dead
        Call GlobalMsg(COLOR_BRIGHTRED & GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker))
        
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Victim) \ 10)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & "You lost no experience points.")
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You received no experience points from that weak insignificant player.")
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & "You lost " & Exp & " experience points.")
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You got " & Exp & " experience points for killing " & GetPlayerName(Victim) & ".")
        End If
        
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER Then
            If TempPlayer(Attacker).Target = Victim Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If
        
        If Room(GetPlayerRoom(Attacker)).Moral <> ROOM_MORAL_ARENA Then
            If GetPlayerPK(Victim) = NO Then
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call GlobalMsg(COLOR_BRIGHTRED & GetPlayerName(Attacker) & " has been deemed a Player Killer!!!")
                End If
            Else
                Call GlobalMsg(COLOR_BRIGHTRED & GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!")
            End If
        End If
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with your bare fists for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit you with their bare fists for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            
            For i = 1 To TotalPlayersOnline
                If GetPlayerRoom(i) = GetPlayerRoom(Attacker) And i <> Attacker And i <> Victim Then
                    Call PlayerMsg(i, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with their bare fists.")
                End If
            Next i
        Else
            Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit you with a " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & " for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
            
            For i = 1 To TotalPlayersOnline
                If GetPlayerRoom(i) = GetPlayerRoom(Attacker) And i <> Attacker And i <> Victim Then
                    Call PlayerMsg(i, COLOR_BRIGHTRED & GetPlayerName(Attacker) & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & " with their " & COLOR_BRIGHTRED & Trim$(Item(n).Name) & COLOR_BRIGHTBLUE & ".")
                End If
            Next i
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Take away 1 stamina
    Call SetPlayerVital(Attacker, Vitals.Stamina, GetPlayerVital(Attacker, Vitals.Stamina) - 1)
    Call SendVital(Attacker, Vitals.Stamina)
    
    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

Public Function FindOpenPlayerSlot() As Long
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Public Function FindOpenRoomItemSlot(ByVal RoomNum As Long) As Long
    Dim i As Long
    
    ' Check for subscript out of range
    If RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_ROOM_ITEMS
        If RoomItem(RoomNum, i).Num = 0 Then
            FindOpenRoomItemSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    Call UpdateHighIndex
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        Call GlobalMsg(COLOR_JOINLEFT & GetPlayerName(Index) & " has joined " & GAME_NAME & "!")
    Else
        Call GlobalMsg(COLOR_WHITE & GetPlayerName(Index) & " has joined " & GAME_NAME & "!")
    End If
    
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' Send an ok to client to start receiving in game data
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 6
    Buffer.WriteInteger SLoginOk
    Buffer.WriteLong Index
    Call SendDataTo(Index, Buffer.ToArray())
    Set Buffer = Nothing
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
        Call SendPlayerExp(Index)
    Next
    Call SendSpells(Index)
    
    ' Send welcome messages
    Call SendWelcome(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerRoom(Index))
    
    ' Send stats for client interface display
    Call SendStats(Index)
    
    ' Send the flag so they know they can start doing stuff
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SInGame
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub LeftGame(ByVal Index As Long)
    Dim n As Long
    
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False
        
        ' Check if player was the only player on the Room and stop Npc processing if so
        If GetTotalRoomPlayers(GetPlayerRoom(Index)) < 1 Then
            PlayersInRoom(GetPlayerRoom(Index)) = NO
        End If
        
        ' Check for boot Room
        If Room(GetPlayerRoom(Index)).BootRoom > 0 Then
            Call SetPlayerRoom(Index, Room(GetPlayerRoom(Index)).BootRoom)
        End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If TempPlayer(Index).InParty = YES Then
            n = TempPlayer(Index).PartyPlayer
            
            Call PlayerMsg(n, COLOR_PINK & GetPlayerName(Index) & " has left " & GAME_NAME & ", disbanning party.")
            TempPlayer(n).InParty = NO
            TempPlayer(n).PartyPlayer = 0
        End If
        
        Call SavePlayer(Index)
        
        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(COLOR_JOINLEFT & GetPlayerName(Index) & " has left " & GAME_NAME & "!")
        Else
            Call GlobalMsg(COLOR_WHITE & GetPlayerName(Index) & " has left " & GAME_NAME & "!")
        End If
        
        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".")
        Call SendLeftGame(Index)
        
        TotalPlayersOnline = TotalPlayersOnline - 1
        Call UpdateHighIndex
        
    End If
    
    Call ClearPlayer(Index)
End Sub

Public Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        ' Make sure we dont try to check a name thats to small
        If Len(GetPlayerName(PlayersOnline(i))) >= Len(Trim$(Name)) Then
            If UCase$(Mid$(GetPlayerName(PlayersOnline(i)), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                FindPlayer = PlayersOnline(i)
                Exit Function
            End If
        End If
    Next
    
    FindPlayer = 0
End Function

Public Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal RoomNum As Long)
    Dim i As Long
    
    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Sub
    End If
    
    ' Find open Room item slot
    i = FindOpenRoomItemSlot(RoomNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, RoomNum)
End Sub

Public Sub SpawnItemSlot(ByVal RoomItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal RoomNum As Long)
    Dim Packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If RoomItemSlot <= 0 Or RoomItemSlot > MAX_ROOM_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Sub
    End If
    
    i = RoomItemSlot
    
    If i <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
                
                RoomItem(RoomNum, i).Num = ItemNum
                RoomItem(RoomNum, i).Value = ItemVal
                
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                        RoomItem(RoomNum, i).Dur = ItemDur
                    Else
                        RoomItem(RoomNum, i).Dur = 0
                    End If
                Else
                    RoomItem(RoomNum, i).Dur = 0
                End If
                
                Set Buffer = New clsBuffer
                
                Buffer.PreAllocate 26 + 4
                Buffer.WriteInteger SSpawnItem
                Buffer.WriteLong RoomNum
                Buffer.WriteLong i
                Buffer.WriteLong ItemNum
                Buffer.WriteLong ItemVal
                Buffer.WriteLong RoomItem(RoomNum, i).Dur
                Call SendDataToAll(Buffer.ToArray())
                
                Set Buffer = Nothing
            End If
        End If
    End If
    
End Sub

Public Sub SpawnAllRoomsItems()
    Dim i As Long
    
    For i = 1 To MAX_ROOMS
        Call SpawnRoomItems(i)
    Next
End Sub

Public Sub SpawnRoomItems(ByVal RoomNum As Long)
Dim i As Long
    ' Check for subscript out of range
    If RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For i = 1 To 5
        If Room(RoomNum).Item(i) > 0 Then
            ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
            If Item(Room(RoomNum).Item(i)).Type = ITEM_TYPE_CURRENCY And Room(RoomNum).ItemVal(i) <= 0 Then
                Call SpawnItem(Room(RoomNum).Item(i), 1, RoomNum)
            ElseIf Item(Room(RoomNum).Item(i)).Type <> ITEM_TYPE_CURRENCY Then
                Call SpawnItem(Room(RoomNum).Item(i), 1, RoomNum)
            Else
                Call SpawnItem(Room(RoomNum).Item(i), Room(RoomNum).ItemVal(i), RoomNum)
            End If
        End If
    Next i
End Sub

Public Sub SpawnNpc(ByVal RoomNpcNum As Long, ByVal RoomNum As Long)
    Dim Packet As String
    Dim NpcNum As Long
    Dim i As Long
    Dim Spawned As Boolean
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If RoomNpcNum <= 0 Or RoomNpcNum > MAX_ROOM_NPCS Or RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Sub
    End If
    
    NpcNum = Room(RoomNum).Npc(RoomNpcNum)
    If NpcNum > 0 Then
        RoomNpc(RoomNum, RoomNpcNum).Num = NpcNum
        RoomNpc(RoomNum, RoomNpcNum).Target = 0
        
        RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
        RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.SP) = GetNpcMaxVital(NpcNum, Vitals.SP)
        
        Spawned = True
        
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            
            Buffer.PreAllocate 12 + 4
            Buffer.WriteInteger SSpawnNpc
            Buffer.WriteLong RoomNum
            Buffer.WriteLong RoomNpcNum
            Buffer.WriteInteger RoomNpc(RoomNum, RoomNpcNum).Num
            
            Call SendDataToAll(Buffer.ToArray())
        End If
    End If
End Sub

Public Sub SpawnRoomNpcs(ByVal RoomNum As Long)
    Dim i As Long
    
    For i = 1 To MAX_ROOM_NPCS
        Call SpawnNpc(i, RoomNum)
    Next
End Sub

Public Sub SpawnAllRoomNpcs()
    Dim i As Long
    
    For i = 1 To MAX_ROOMS
        Call SpawnRoomNpcs(i)
    Next
End Sub

Public Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim zType As Long
    ' Check attack timer
    If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function
    
    ' Make sure they are on the same Room
    If Not GetPlayerRoom(Attacker) = GetPlayerRoom(Victim) Then Exit Function
    
    ' Make sure we dont attack the player if they are switching Rooms
    If TempPlayer(Victim).GettingRoom = YES Then Exit Function
    
    ' For debugging
    zType = Room(GetPlayerRoom(Attacker)).Moral
    ' Check if Room is attackable
    If (Not Room(GetPlayerRoom(Attacker)).Moral = ROOM_MORAL_NONE) And (Not Room(GetPlayerRoom(Attacker)).Moral = ROOM_MORAL_ARENA) Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, COLOR_BRIGHTRED & "This is a safe zone!")
            Exit Function
        End If
    End If
    
    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    ' Make sure they have stamina
    If GetPlayerVital(Attacker, Vitals.Stamina) <= 0 Then
        Call PlayerMsg(Attacker, COLOR_BRIGHTRED & "You are exhausted!")
        Exit Function
    End If
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You cannot attack any player for thou art an admin!")
        Exit Function
    End If
    
    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, COLOR_BRIGHTRED & "You cannot attack " & GetPlayerName(Victim) & "!")
        Exit Function
    End If
    
'    ' Make sure attacker is high enough level
'    If GetPlayerLevel(Attacker) < 5 Then
'        Call PlayerMsg(Attacker, COLOR_BRIGHTRED & "You are below level 5, you cannot attack another player yet!")
'        Exit Function
'    End If
'
'    ' Make sure victim is high enough level
'    If GetPlayerLevel(Victim) < 5 Then
'        Call PlayerMsg(Attacker, COLOR_BRIGHTRED & GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!")
'        Exit Function
'    End If
    
    CanAttackPlayer = True
    
End Function

Public Function CanAttackNpc(ByVal Attacker As Long, ByVal RoomNpcNum As Long) As Boolean
    Dim RoomNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or RoomNpcNum <= 0 Or RoomNpcNum > MAX_ROOM_NPCS Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If RoomNpc(GetPlayerRoom(Attacker), RoomNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    RoomNum = GetPlayerRoom(Attacker)
    NpcNum = RoomNpc(RoomNum, RoomNpcNum).Num
    
    ' Make sure the Npc isn't already dead
    If RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure they have stamina
    If GetPlayerVital(Attacker, Vitals.Stamina) <= 0 Then
        Call PlayerMsg(Attacker, COLOR_BRIGHTRED & "You are exhausted!")
        Exit Function
    End If
    
    ' Make sure they are on the same Room
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + 1000 Then
'            ' Check if at same coordinates
'            Select Case GetPlayerDir(Attacker)
'            Case DIR_NORTH
'                NpcX = RoomNpc(RoomNum, RoomNpcNum).X
'                NpcY = RoomNpc(RoomNum, RoomNpcNum).y + 1
'            Case DIR_SOUTH
'                NpcX = RoomNpc(RoomNum, RoomNpcNum).X
'                NpcY = RoomNpc(RoomNum, RoomNpcNum).y - 1
'            Case DIR_WEST
'                NpcX = RoomNpc(RoomNum, RoomNpcNum).X + 1
'                NpcY = RoomNpc(RoomNum, RoomNpcNum).y
'            Case DIR_EAST
'                NpcX = RoomNpc(RoomNum, RoomNpcNum).X - 1
'                NpcY = RoomNpc(RoomNum, RoomNpcNum).y
'            End Select
            
'            If NpcX = GetPlayerX(Attacker) Then
'                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(NpcNum).Behavior <> Npc_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> Npc_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, COLOR_BRIGHTBLUE & "You cannot attack a " & Trim$(Npc(NpcNum).Name) & "!")
                    End If
'                End If
'            End If
        End If
    End If
End Function

Public Function CanNpcAttackPlayer(ByVal RoomNpcNum As Long, ByVal Index As Long) As Boolean
    Dim RoomNum As Long
    Dim NpcNum As Long
    
    ' Check for subscript out of range
    If RoomNpcNum <= 0 Or RoomNpcNum > MAX_ROOM_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If RoomNpc(GetPlayerRoom(Index), RoomNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    RoomNum = GetPlayerRoom(Index)
    NpcNum = RoomNpc(RoomNum, RoomNpcNum).Num
    
    ' Make sure the Npc isn't already dead
    If RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure Npcs dont attack more then once every 3 second
    If GetTickCount < RoomNpc(RoomNum, RoomNpcNum).AttackTimer + 3000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching Rooms
    If TempPlayer(Index).GettingRoom = YES Then
        Exit Function
    End If
    
    RoomNpc(RoomNum, RoomNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same Room
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Can attack player
            CanNpcAttackPlayer = True
        End If
    End If
End Function

Public Sub NpcAttackPlayer(ByVal RoomNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim RoomNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If RoomNpcNum <= 0 Or RoomNpcNum > MAX_ROOM_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If RoomNpc(GetPlayerRoom(Victim), RoomNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    RoomNum = GetPlayerRoom(Victim)
    Name = Trim$(Npc(RoomNpc(RoomNum, RoomNpcNum).Num).Name)
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 6
    Buffer.WriteInteger SNpcAttack
    Buffer.WriteLong RoomNpcNum
    Call SendDataToRoom(RoomNum, Buffer.ToArray())
    
    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        Call PlayerMsg(Victim, COLOR_BRIGHTBLUE & "A " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " hit you for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
        Call RoomMsgBut(RoomNum, Victim, COLOR_BRIGHTBLUE & "A " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & ".")
        
        ' Player is dead
        Call GlobalMsg(COLOR_BRIGHTRED & GetPlayerName(Victim) & " has been killed by a " & Name)
        
        ' Calculate exp to give attacker
        Exp = GetPlayerExp(Victim) \ 3
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & "You lost no experience points.")
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, COLOR_BRIGHTRED & "You lost " & Exp & " experience points.")
        End If
        
        ' Set Npc target to 0
        RoomNpc(RoomNum, RoomNpcNum).Target = 0
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Say damage
        Call PlayerMsg(Victim, COLOR_BRIGHTBLUE & "A " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " hit you for " & COLOR_BRIGHTRED & Damage & COLOR_BRIGHTBLUE & " hit points.")
        Call RoomMsgBut(RoomNum, Victim, COLOR_BRIGHTBLUE & "A " & COLOR_BRIGHTRED & Name & COLOR_BRIGHTBLUE & " hit " & COLOR_BRIGHTRED & GetPlayerName(Victim) & COLOR_BRIGHTBLUE & ".")
    End If
End Sub

Public Function GetTotalRoomPlayers(ByVal RoomNum As Long) As Long
    Dim i As Long
    Dim n As Long
    
    n = 0
    
    For i = 1 To High_Index
        If IsPlaying(i) Then
            If GetPlayerRoom(i) = RoomNum Then
                n = n + 1
            End If
        End If
    Next
    
    GetTotalRoomPlayers = n
End Function

Public Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long
    Dim y As Long
    
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    Select Case Vital
    Case HP
        X = Npc(NpcNum).Stat(Stats.Strength)
        y = Npc(NpcNum).Stat(Stats.Defense)
        GetNpcMaxVital = X * y
    Case MP
        GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Magic) * 2
    Case SP
        GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Speed) * 2
    End Select
End Function

Public Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long
    
    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
    Case HP
        i = Npc(NpcNum).Stat(Stats.Defense) \ 3
        If i < 1 Then i = 1
        GetNpcVitalRegen = i
        'Case MP
        
        'Case SP
        
    End Select
End Function

Public Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim WeaponSlot As Long
    
    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > High_Index Then
        Exit Function
    End If
    
    GetPlayerDamage = (GetPlayerStat(Index, Stats.Strength) \ 2)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
        WeaponSlot = GetPlayerEquipmentSlot(Index, Weapon)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
    End If
End Function

Public Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim ArmorSlot As Long
    Dim HelmSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > High_Index Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerEquipmentSlot(Index, Armor)
    HelmSlot = GetPlayerEquipmentSlot(Index, Helmet)
    
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Defense) \ 5)
    
    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
    End If
End Function

Public Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Public Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    
    ShieldSlot = GetPlayerEquipmentSlot(Index, Shield)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Defense) \ 2) + (GetPlayerLevel(Index) \ 2)
            
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Public Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellNum As Long
    Dim MPReq As Long
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim Casted As Boolean
    Dim CanCast As Boolean
    Dim TargetType As Byte
    Dim TargetName As String
    Dim Buffer As clsBuffer
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "You do not have this spell!")
        Exit Sub
    End If
    
    ' (does not check for level requirement)
    ' Make sure they are the right level
    'If ?? > GetPlayerLevel(Index) Then
    '    Call PlayerMsg(Index, COLOR_BRIGHTRED & "You must be level " & ??? & " to cast this spell.")
    '    Exit Sub
    'End If
    
    MPReq = Spell(SpellNum).MPReq
    
    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPReq Then
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Not enough mana points!")
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < TempPlayer(Index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' *** Self Cast Spells ***
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
        
        If n > 0 Then
            Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
            Call RoomMsg(GetPlayerRoom(Index), COLOR_BRIGHTCYAN & GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & ".")
            
            ' Take away the mana points
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPReq)
            Call SendVital(Index, Vitals.MP)
            Casted = True
        Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your inventory is full!")
        End If
        
        Exit Sub
    End If
    
    n = TempPlayer(Index).Target
    TargetType = TempPlayer(Index).TargetType
    
    Select Case TargetType
    Case TARGET_TYPE_PLAYER
        
        If IsPlaying(n) Then
            
            If GetPlayerVital(n, Vitals.HP) > 0 Then
                If GetPlayerRoom(Index) = GetPlayerRoom(n) Then
                    'If GetPlayerLevel(Index) >= 10 Then
                    'If GetPlayerLevel(n) >= 10 Then
                    If (Room(GetPlayerRoom(Index)).Moral = ROOM_MORAL_NONE) Or (Room(GetPlayerRoom(Index)).Moral = ROOM_MORAL_ARENA) Then
                        If GetPlayerAccess(Index) <= 0 Then
                            If GetPlayerAccess(n) <= 0 Then
                                If n <> Index Then
                                    CanCast = True
                                End If
                            End If
                        End If
                    End If
                    'End If
                    'End If
                End If
            End If
            
            TargetName = GetPlayerName(n)
            
            If Spell(SpellNum).Type = SPELL_TYPE_SUBHP Or Spell(SpellNum).Type = SPELL_TYPE_SUBMP Or Spell(SpellNum).Type = SPELL_TYPE_SUBSP Then
            
                If CanCast Then
                    Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_SUBHP
                        Damage = (GetPlayerStat(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - GetPlayerProtection(n)
                        If Damage > 0 Then
                            Call AttackPlayer(Index, n, Damage)
                        Else
                            Call PlayerMsg(Index, COLOR_BRIGHTRED & "The spell was to weak to hurt " & GetPlayerName(n) & "!")
                        End If
                        
                    Case SPELL_TYPE_SUBMP
                        Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) - Spell(SpellNum).Data1)
                        Call SendVital(n, Vitals.MP)
                        
                    Case SPELL_TYPE_SUBSP
                        Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) - Spell(SpellNum).Data1)
                        Call SendVital(n, Vitals.SP)
                    End Select
                    
                    Casted = True
                    
                End If
            
            ElseIf Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
        
                If GetPlayerRoom(Index) = GetPlayerRoom(n) Then
                    CanCast = True
                End If
                
                If CanCast Then
                    Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_ADDHP
                        Call SetPlayerVital(n, Vitals.HP, GetPlayerVital(n, Vitals.HP) + Spell(SpellNum).Data1)
                        Call SendVital(n, Vitals.HP)
                        
                    Case SPELL_TYPE_ADDMP
                        Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) + Spell(SpellNum).Data1)
                        Call SendVital(n, Vitals.MP)
                        
                    Case SPELL_TYPE_ADDSP
                        Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) + Spell(SpellNum).Data1)
                        Call SendVital(n, Vitals.SP)
                    End Select
                    
                    Casted = True
                End If
        
            End If
        End If

    Case TARGET_TYPE_NPC
        
        If Npc(RoomNpc(GetPlayerRoom(Index), n).Num).Behavior <> Npc_BEHAVIOR_FRIENDLY Then
            If Npc(RoomNpc(GetPlayerRoom(Index), n).Num).Behavior <> Npc_BEHAVIOR_SHOPKEEPER Then
                CanCast = True
            End If
        End If
        
        TargetName = Npc(RoomNpc(GetPlayerRoom(Index), n).Num).Name
        
        If CanCast Then
            Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_ADDHP
                RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.HP) = RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.HP) + Spell(SpellNum).Data1
                
            Case SPELL_TYPE_SUBHP
                
                Damage = (GetPlayerStat(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - (Npc(RoomNpc(GetPlayerRoom(Index), n).Num).Stat(Stats.Defense) \ 2)
                If Damage > 0 Then
                    Call AttackNpc(Index, n, Damage)
                Else
                    Call PlayerMsg(Index, COLOR_BRIGHTRED & "The spell was to weak to hurt " & Trim$(Npc(RoomNpc(GetPlayerRoom(Index), n).Num).Name) & "!")
                End If
                
            Case SPELL_TYPE_ADDMP
                RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.MP) = RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.MP) + Spell(SpellNum).Data1
                
            Case SPELL_TYPE_SUBMP
                RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.MP) = RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.MP) - Spell(SpellNum).Data1
                
            Case SPELL_TYPE_ADDSP
                RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.SP) = RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.SP) + Spell(SpellNum).Data1
                
            Case SPELL_TYPE_SUBSP
                RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.SP) = RoomNpc(GetPlayerRoom(Index), n).Vital(Vitals.SP) - Spell(SpellNum).Data1
            End Select
            
            Casted = True
        End If
        
    End Select

    If Casted Then
        Call RoomMsg(GetPlayerRoom(Index), COLOR_BRIGHTCYAN & GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & Trim$(TargetName) & ".")
        
        Set Buffer = New clsBuffer
        Buffer.PreAllocate 11
        Buffer.WriteInteger SCastSpell
        Buffer.WriteByte TargetType
        Buffer.WriteLong n
        Buffer.WriteLong SpellNum
        Call SendDataToRoom(GetPlayerRoom(Index), Buffer.ToArray())
        
        ' Take away the mana points
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPReq)
        Call SendVital(Index, Vitals.MP)
        
        TempPlayer(Index).AttackTimer = GetTickCount
        TempPlayer(Index).CastedSpell = YES
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Could not cast spell!")
    End If

End Sub

Public Sub PlayerWarp(ByVal Index As Long, ByVal RoomNum As Long)
    Dim ShopNum As Long
    Dim OldRoom As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or RoomNum <= 0 Or RoomNum > MAX_ROOMS Then
        Exit Sub
    End If
    
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    
    ' Check if there was a shop on the Room the player is leaving, and if so say goodbye
    ShopNum = Room(GetPlayerRoom(Index)).Shop
    If ShopNum > 0 Then
        If LenB(Trim$(Shop(ShopNum).LeaveSay)) > 0 Then
            Call PlayerMsg(Index, COLOR_SAY & Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).LeaveSay) & "'")
        End If
    End If
    
    ' Save old Room to send erase player data to
    OldRoom = GetPlayerRoom(Index)
    
    If OldRoom <> RoomNum Then
        'Call SendLeaveRoom(Index, OldRoom)
    End If
    
    Call SetPlayerRoom(Index, RoomNum)
    
    ' Check if there is a shop on the Room and say hello if so
    ShopNum = Room(GetPlayerRoom(Index)).Shop
    If ShopNum > 0 Then
        If LenB(Trim$(Shop(ShopNum).JoinSay)) > 0 Then
            Call PlayerMsg(Index, COLOR_SAY & Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).JoinSay) & "'")
        End If
    End If
    
'    ROOM_MORAL_NONE = 0
'    ROOM_MORAL_SAFE = 1
'    ROOM_MORAL_INN = 2
'    ROOM_MORAL_ARENA = 3

    ' Send the room description
    Select Case Room(RoomNum).Moral
        Case ROOM_MORAL_NONE
            Call PlayerMsg(Index, COLOR_YELLOW & "<< " & COLOR_BRIGHTRED & Trim$(Room(RoomNum).Name) & COLOR_YELLOW & " >>")
        Case ROOM_MORAL_SAFE
            Call PlayerMsg(Index, COLOR_YELLOW & "<< " & COLOR_BRIGHTCYAN & Trim$(Room(RoomNum).Name) & COLOR_YELLOW & " >>")
        Case ROOM_MORAL_INN
            Call PlayerMsg(Index, COLOR_YELLOW & "<< " & COLOR_WHITE & Trim$(Room(RoomNum).Name) & COLOR_YELLOW & " >>")
        Case ROOM_MORAL_ARENA
            Call PlayerMsg(Index, COLOR_YELLOW & "<< " & COLOR_BRIGHTBLUE & Trim$(Room(RoomNum).Name) & COLOR_YELLOW & " >>")
    End Select
    
    Call PlayerMsg(Index, COLOR_BRIGHTBLUE & Trim$(Room(RoomNum).sDesc))
    Call PlayerMsg(Index, COLOR_BRIGHTBLUE & Trim$(Room(RoomNum).eDesc))
    
    ' Now we check if there were any players left on the Room the player just left, and if not stop processing Npcs
    If GetTotalRoomPlayers(OldRoom) = 0 Then
        PlayersInRoom(OldRoom) = NO
        
        ' Regenerate all Npcs' health
        'For i = 1 To MAX_Room_NPCS
        '    If RoomNpc(OldRoom, i).Num > 0 Then
        '        'RoomNpc(OldRoom, i).Vital(Vitals.HP) = GetNpcMaxVital(RoomNpc(OldRoom, i).Num, Vitals.HP)
        '    End If
        'Next
        
    End If
    
    ' Sets it so we know to process Npcs on the Room
    PlayersInRoom(RoomNum) = YES
    
    TempPlayer(Index).GettingRoom = YES
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 10
    Buffer.WriteInteger SCheckForRoom
    Buffer.WriteLong RoomNum
    Buffer.WriteLong Room(RoomNum).Revision
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long)
    Dim RoomNum As Long
    Dim NewRoom As Long
    Dim Moved As Byte
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_NORTH Or Dir > DIR_EAST Then
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    
    Moved = NO
    RoomNum = GetPlayerRoom(Index)
    
    Select Case Dir
    Case DIR_NORTH
        ' Check to see if we can move them to the another Room
        If Room(GetPlayerRoom(Index)).North > 0 Then
            NewRoom = Room(GetPlayerRoom(Index)).North
            Call PlayerMsg(Index, COLOR_GREEN & "You moved North.")
            Call RoomMsgBut(RoomNum, Index, COLOR_GREEN & GetPlayerName(Index) & " moved North.")
            Call RoomMsg(NewRoom, COLOR_YELLOW & GetPlayerName(Index) & " arrived from the South.")
            Call PlayerWarp(Index, Room(GetPlayerRoom(Index)).North)
            Moved = YES
        Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You cannot go that way.")
        End If
        
    Case DIR_SOUTH
        ' Check to see if we can move them to the another Room
        If Room(GetPlayerRoom(Index)).South > 0 Then
            NewRoom = Room(GetPlayerRoom(Index)).South
            Call PlayerMsg(Index, COLOR_GREEN & "You moved South.")
            Call RoomMsgBut(RoomNum, Index, COLOR_GREEN & GetPlayerName(Index) & " moved South.")
            Call RoomMsg(NewRoom, COLOR_YELLOW & GetPlayerName(Index) & " arrived from the North.")
            Call PlayerWarp(Index, Room(GetPlayerRoom(Index)).South)
            Moved = YES
        Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You cannot go that way.")
        End If
        
    Case DIR_WEST
        ' Check to see if we can move them to the another Room
        If Room(GetPlayerRoom(Index)).West > 0 Then
            NewRoom = Room(GetPlayerRoom(Index)).West
            Call PlayerMsg(Index, COLOR_GREEN & "You moved West.")
            Call RoomMsgBut(RoomNum, Index, COLOR_GREEN & GetPlayerName(Index) & " moved West.")
            Call RoomMsg(NewRoom, COLOR_YELLOW & GetPlayerName(Index) & " arrived from the East.")
            Call PlayerWarp(Index, Room(GetPlayerRoom(Index)).West)
            Moved = YES
        Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You cannot go that way.")
        End If
        
    Case DIR_EAST
        ' Check to see if we can move them to the another Room
        If Room(GetPlayerRoom(Index)).East > 0 Then
            NewRoom = Room(GetPlayerRoom(Index)).East
            Call PlayerMsg(Index, COLOR_GREEN & "You moved East.")
            Call RoomMsgBut(RoomNum, Index, COLOR_GREEN & GetPlayerName(Index) & " moved East.")
            Call RoomMsg(NewRoom, COLOR_YELLOW & GetPlayerName(Index) & " arrived from the West.")
            Call PlayerWarp(Index, Room(GetPlayerRoom(Index)).East)
            Moved = YES
        Else
            Call PlayerMsg(Index, COLOR_BRIGHTRED & "You cannot go that way.")
        End If
    End Select
End Sub

Private Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        Slot = GetPlayerEquipmentSlot(Index, i)
        If Slot > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, Slot)
            
            If ItemNum > 0 Then
                Select Case i
                Case Equipment.Weapon
                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipmentSlot Index, 0, i
                Case Equipment.Armor
                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipmentSlot Index, 0, i
                Case Equipment.Helmet
                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipmentSlot Index, 0, i
                Case Equipment.Shield
                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipmentSlot Index, 0, i
                End Select
            Else
                SetPlayerEquipmentSlot Index, 0, i
            End If
        End If
    Next
End Sub

Public Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Public Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Public Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long
    Dim n As Long
    Dim TakeItem As Boolean
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, i)).Type
                Case ITEM_TYPE_WEAPON
                    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                        If i = GetPlayerEquipmentSlot(Index, Weapon) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                            Call SendWornEquipment(Index)
                            TakeItem = True
                        Else
                            ' Check if the item we are taking isn't already equipped
                            If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon)) Then
                                TakeItem = True
                            End If
                        End If
                    Else
                        TakeItem = True
                    End If
                    
                Case ITEM_TYPE_ARMOR
                    If GetPlayerEquipmentSlot(Index, Armor) > 0 Then
                        If i = GetPlayerEquipmentSlot(Index, Armor) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Armor)
                            Call SendWornEquipment(Index)
                            TakeItem = True
                        Else
                            ' Check if the item we are taking isn't already equipped
                            If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor)) Then
                                TakeItem = True
                            End If
                        End If
                    Else
                        TakeItem = True
                    End If
                    
                Case ITEM_TYPE_HELMET
                    If GetPlayerEquipmentSlot(Index, Helmet) > 0 Then
                        If i = GetPlayerEquipmentSlot(Index, Helmet) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                            Call SendWornEquipment(Index)
                            TakeItem = True
                        Else
                            ' Check if the item we are taking isn't already equipped
                            If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet)) Then
                                TakeItem = True
                            End If
                        End If
                    Else
                        TakeItem = True
                    End If
                    
                Case ITEM_TYPE_SHIELD
                    If GetPlayerEquipmentSlot(Index, Shield) > 0 Then
                        If i = GetPlayerEquipmentSlot(Index, Shield) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Shield)
                            Call SendWornEquipment(Index)
                            TakeItem = True
                        Else
                            ' Check if the item we are taking isn't already equipped
                            If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield)) Then
                                TakeItem = True
                            End If
                        End If
                    Else
                        TakeItem = True
                    End If
                End Select
                
                
                n = Item(GetPlayerInvItemNum(Index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If
            
            If TakeItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next
End Sub

Public Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your inventory is full.")
    End If
End Sub

Public Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Public Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub PlayerRoomGetItem(ByVal Index As Long, ByVal mItem As Long)
    Dim i As Long
    Dim n As Long
    Dim RoomNum As Long
    Dim Msg As String
    
    If Not IsPlaying(Index) Then Exit Sub
    
    RoomNum = GetPlayerRoom(Index)
    
    ' See if theres even an item here
    If (RoomItem(RoomNum, mItem).Num > 0) Then
        If (RoomItem(RoomNum, mItem).Num <= MAX_ITEMS) Then
                    
            ' Find open slot
            n = FindOpenInvSlot(Index, RoomItem(RoomNum, mItem).Num)
            
            ' Open slot available?
            If n <> 0 Then
                ' Set item in players inventor
                Call SetPlayerInvItemNum(Index, n, RoomItem(RoomNum, mItem).Num)
                If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                    Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + RoomItem(RoomNum, mItem).Value)
                    Msg = "You picked up " & RoomItem(RoomNum, mItem).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                Else
                    Call SetPlayerInvItemValue(Index, n, 0)
                    Msg = "You picked up a " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                End If
                Call SetPlayerInvItemDur(Index, n, RoomItem(RoomNum, mItem).Dur)
                
                ' Erase item from the Room
                RoomItem(RoomNum, mItem).Num = 0
                RoomItem(RoomNum, mItem).Value = 0
                RoomItem(RoomNum, mItem).Dur = 0
                RoomItem(RoomNum, mItem).X = 0
                RoomItem(RoomNum, mItem).y = 0
                
                Call SendInventoryUpdate(Index, n)
                Call SpawnItemSlot(mItem, 0, 0, 0, GetPlayerRoom(Index))
                Call PlayerMsg(Index, COLOR_YELLOW & Msg)
            Else
                Call PlayerMsg(Index, COLOR_BRIGHTRED & "Your inventory is full.")
            End If
            
        End If
    End If
End Sub

Public Sub PlayerRoomDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            
            i = FindOpenRoomItemSlot(GetPlayerRoom(Index))
            
            If i <> 0 Then
                RoomItem(GetPlayerRoom(Index), i).Dur = 0
                
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerEquipmentSlot(Index, Armor) Then
                        Call SetPlayerEquipmentSlot(Index, 0, Armor)
                        Call SendWornEquipment(Index)
                    End If
                    RoomItem(GetPlayerRoom(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerEquipmentSlot(Index, Weapon) Then
                        Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                        Call SendWornEquipment(Index)
                    End If
                    RoomItem(GetPlayerRoom(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerEquipmentSlot(Index, Helmet) Then
                        Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                        Call SendWornEquipment(Index)
                    End If
                    RoomItem(GetPlayerRoom(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerEquipmentSlot(Index, Shield) Then
                        Call SetPlayerEquipmentSlot(Index, 0, Shield)
                        Call SendWornEquipment(Index)
                    End If
                    RoomItem(GetPlayerRoom(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
                
                RoomItem(GetPlayerRoom(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        RoomItem(GetPlayerRoom(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call RoomMsg(GetPlayerRoom(Index), COLOR_YELLOW & GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".")
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        RoomItem(GetPlayerRoom(Index), i).Value = Amount
                        Call RoomMsg(GetPlayerRoom(Index), COLOR_YELLOW & GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".")
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    RoomItem(GetPlayerRoom(Index), i).Value = 0
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                        Call RoomMsg(GetPlayerRoom(Index), COLOR_YELLOW & GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".")
                    Else
                        Call RoomMsg(GetPlayerRoom(Index), COLOR_YELLOW & GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".")
                    End If
                    
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                End If
                
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free Room item slot
                Call SpawnItemSlot(i, RoomItem(GetPlayerRoom(Index), i).Num, Amount, RoomItem(GetPlayerRoom(Index), i).Dur, GetPlayerRoom(Index))
            Else
                Call PlayerMsg(Index, COLOR_BRIGHTRED & "To many items already on the ground.")
            End If
        End If
    End If
End Sub

Public Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollOver As Long
    
    ' Check if attacker got a level up
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        expRollOver = CLng(GetPlayerExp(Index) - GetPlayerNextLevel(Index))
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
        
        ' Get the amount of skill points to add
        i = (GetPlayerStat(Index, Stats.Speed) \ 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
        
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
        Call SetPlayerExp(Index, expRollOver)
        Call GlobalMsg(COLOR_BROWN & GetPlayerName(Index) & " has gained a level!")
        Call PlayerMsg(Index, COLOR_BRIGHTBLUE & "You have gained a level!  You now have " & GetPlayerPOINTS(Index) & " stat points to distribute.")
        Call SendStats(Index)
    End If
    
End Sub

Public Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long
    
    ' Prevent subscript out of range
    If Not IsPlaying(Index) Or Index <= 0 Or Index > High_Index Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
    Case HP
        i = (GetPlayerStat(Index, Stats.Defense) \ 2)
    Case MP
        i = (GetPlayerStat(Index, Stats.Magic) \ 2)
    Case SP
        i = (GetPlayerStat(Index, Stats.Speed) \ 2)
    End Select
    
    If i < 2 Then i = 2
    
    GetPlayerVitalRegen = i
End Function

' ToDo
Public Sub OnDeath(ByVal Index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    
    ' Drop all worn items
    If Room(GetPlayerRoom(Index)).Moral <> ROOM_MORAL_ARENA Then
        For i = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipmentSlot(Index, i) > 0 Then
                PlayerRoomDropItem Index, GetPlayerEquipmentSlot(Index, i), 0
            End If
        Next
    End If
    
    ' Warp player away
    Call PlayerWarp(Index, START_ROOM)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SetPlayerVital(Index, Vitals.SP, GetPlayerMaxVital(Index, Vitals.SP))
    Call SetPlayerVital(Index, Vitals.Stamina, GetPlayerMaxVital(Index, Vitals.Stamina))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    Call SendVital(Index, Vitals.SP)
    Call SendVital(Index, Vitals.Stamina)
    
    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If
    
End Sub

Public Sub DamageEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment)
    Dim Slot As Long
    
    Slot = GetPlayerEquipmentSlot(Index, EquipmentSlot)
    
    If Slot > 0 Then
        Call SetPlayerInvItemDur(Index, Slot, GetPlayerInvItemDur(Index, Slot) - 1)
        
        If GetPlayerInvItemDur(Index, Slot) <= 0 Then
            Call PlayerMsg(Index, COLOR_YELLOW & "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " has broken.")
            Call TakeItem(Index, GetPlayerInvItemNum(Index, Slot), 0)
        Else
            If GetPlayerInvItemDur(Index, Slot) <= 5 Then
                Call PlayerMsg(Index, COLOR_YELLOW & "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " is about to break!")
            End If
        End If
    End If
End Sub

Public Sub UpdateHighIndex()
    Dim i As Integer
    Dim array_index As Integer
    Dim Buffer As clsBuffer
    
    ' no players are logged in
    If TotalPlayersOnline < 1 Then
        High_Index = 0
        Exit Sub
    End If
    
    ' new size
    ReDim PlayersOnline(1 To TotalPlayersOnline)
    
    For i = 1 To MAX_PLAYERS
        If LenB((GetPlayerLogin(i))) > 0 Then
            High_Index = i
            array_index = array_index + 1
            PlayersOnline(array_index) = i
            
            ' early finish if all players are found
            If array_index >= TotalPlayersOnline Then
                Exit For
            End If
            
        End If
    Next
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 6
    Buffer.WriteInteger SHighIndex
    Buffer.WriteLong High_Index
    Call SendDataToAll(Buffer.ToArray())
End Sub

