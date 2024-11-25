Attribute VB_Name = "modServerLoop"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' halts thread of execution
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public ServerOnline As Boolean ' Used for server loop
Private GiveNpcHPTimer As Long  ' Used for Npc HP regeneration

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Private Secs As Long

Public Sub ServerLoop()
    Dim i As Long
    Dim X As Long
    Dim y As Long
    
    Dim Tick As Long
    
    Dim tmr500 As Long
    Dim tmr1000 As Long
    
    Dim LastUpdateSavePlayers As Long
    Dim LastUpdateRoomSpawnItems As Long
    Dim LastUpdatePlayerVitals As Long
    
    Dim Buffer As clsBuffer
    
    Do While ServerOnline
        Tick = GetTickCount
        
        If Tick > tmr500 Then
            ' Check for disconnections
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            ' Process Npc AI
            UpdateNpcAI
            
            tmr500 = GetTickCount + 500
        End If
        
        If Tick > tmr1000 Then
            ' Handle shutting down server
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            tmr1000 = GetTickCount + 1000
        End If
        
        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If
        
        ' Checks to spawn Room items every 5 minutes - Can be tweaked
        If Tick > LastUpdateRoomSpawnItems Then
            UpdateRoomSpawnItems
            LastUpdateRoomSpawnItems = GetTickCount + 300000
        End If
        
        ' Checks to save players every 10 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 600000
        End If
        
        Sleep 1
        DoEvents
        
    Loop
    
End Sub

Private Sub UpdateRoomSpawnItems()
    Dim X As Long
    Dim y As Long
    
    ' This is used for respawning Room items
    For y = 1 To MAX_ROOMS
        ' Make sure no one is on the Room when it respawns
        If Not PlayersInRoom(y) Then
            ' Clear out unnecessary junk
            For X = 1 To MAX_ROOM_ITEMS
                Call ClearRoomItem(X, y)
            Next
            
            ' Spawn the items
            Call SpawnRoomItems(y)
            Call SendRoomItemsToAll(y)
        End If
        DoEvents
    Next
    
End Sub

Private Sub UpdateNpcAI()
    Dim i As Long, n As Long
    Dim RoomNum As Long, RoomNpcNum As Long
    Dim NpcNum As Long, Target As Long
    Dim TickCount As Long
    Dim Damage As Long
    Dim DistanceX As Long
    Dim DistanceY As Long
    Dim DidWalk As Boolean
    
    For RoomNum = 1 To MAX_ROOMS
        If True Then
            TickCount = GetTickCount
            
            For RoomNpcNum = 1 To MAX_ROOM_NPCS
                NpcNum = RoomNpc(RoomNum, RoomNpcNum).Num
                
                ' Make sure theres a Npc with the Room
                If NpcNum > 0 Then
                    
                    ' Get the target
                    Target = RoomNpc(RoomNum, RoomNpcNum).Target
                    
                    ' /////////////////////////////////////////
                    ' // This is used for ATTACKING ON SIGHT //
                    ' /////////////////////////////////////////
                    ' If the Npc is a attack on sight, search for a player on the Room
                    If Npc(NpcNum).Behavior = Npc_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = Npc_BEHAVIOR_GUARD Then
                        ' First check if they don't have a target before looping...
                        If Target = 0 Then
                            For i = 1 To High_Index
                                If IsPlaying(i) Then
                                    If GetPlayerRoom(i) = RoomNum Then
                                        If GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                            n = Npc(NpcNum).Range
                                            
                                            If Npc(NpcNum).Behavior = Npc_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                    Call PlayerMsg(i, COLOR_SAY & "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.")
                                                End If
                                                
                                                RoomNpc(RoomNum, RoomNpcNum).Target = i
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                    
                    
                    ' /////////////////////////////////////////////
                    ' // This is used for Npc walking/targetting //
                    ' /////////////////////////////////////////////
                    ' Check to see if its time for the Npc to walk
                    If Npc(NpcNum).Behavior <> Npc_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) Then
                                If GetPlayerRoom(Target) = RoomNum Then
                                    ' /////////////////////////////////////////////
                                    ' // This is used for Npcs to attack players //
                                    ' /////////////////////////////////////////////
                                    ' Can the Npc attack the player?
                                    If CanNpcAttackPlayer(RoomNpcNum, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = Npc(NpcNum).Stat(Stats.Strength) - GetPlayerProtection(Target)
                                            Call NpcAttackPlayer(RoomNpcNum, Target, Damage)
                                        Else
                                            Call PlayerMsg(Target, COLOR_BRIGHTCYAN & "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerEquipmentSlot(Target, Shield))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!")
                                        End If
                                    End If
                                Else
                                    RoomNpc(RoomNum, RoomNpcNum).Target = 0
                                End If
                            Else
                                RoomNpc(RoomNum, RoomNpcNum).Target = 0
                            End If
                        End If
                    End If
                    
                    ' ////////////////////////////////////////////
                    ' // This is used for regenerating Npc's HP //
                    ' ////////////////////////////////////////////
                    If TickCount > GiveNpcHPTimer + 10000 Then
                        If RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) > 0 Then
                            RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) = RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                            
                            ' Check if they have more then they should and if so just set it to max
                            If RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                RoomNpc(RoomNum, RoomNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                            End If
                        End If
                        GiveNpcHPTimer = TickCount
                    End If
                    
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an Npc //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an Npc or not
                If RoomNpc(RoomNum, RoomNpcNum).Num = 0 Then
                    If Room(RoomNum).Npc(RoomNpcNum) > 0 Then
                        If TickCount > RoomNpc(RoomNum, RoomNpcNum).SpawnWait + (Npc(Room(RoomNum).Npc(RoomNpcNum)).SpawnSecs * 1000) Then
                            Call SpawnNpc(RoomNpcNum, RoomNum)
                        End If
                    End If
                End If
                
            Next
        End If
        DoEvents
    Next
    
End Sub

Private Sub UpdatePlayerVitals()
    Dim i As Long
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerVital(PlayersOnline(i), Vitals.HP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.HP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.HP, GetPlayerVital(PlayersOnline(i), Vitals.HP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.HP))
            Call SendVital(PlayersOnline(i), Vitals.HP)
        End If
        If GetPlayerVital(PlayersOnline(i), Vitals.MP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.MP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.MP, GetPlayerVital(PlayersOnline(i), Vitals.MP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.MP))
            Call SendVital(PlayersOnline(i), Vitals.MP)
        End If
        If GetPlayerVital(PlayersOnline(i), Vitals.SP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.SP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.SP, GetPlayerVital(PlayersOnline(i), Vitals.SP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.SP))
            Call SendVital(PlayersOnline(i), Vitals.SP)
        End If
        If GetPlayerVital(PlayersOnline(i), Vitals.Stamina) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.Stamina) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.Stamina, GetPlayerVital(PlayersOnline(i), Vitals.Stamina) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.Stamina))
            Call SendVital(PlayersOnline(i), Vitals.Stamina)
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long
    
    If TotalPlayersOnline > 0 Then
        Call TextAdd("Saving all online players...")
        'Call AdminMsg(COLOR_PINK & "Saving all online players...")
        
        For i = 1 To TotalPlayersOnline
            Call SavePlayer(PlayersOnline(i))
            DoEvents
        Next
    End If
    
End Sub

Private Sub HandleShutdown()
    If Secs <= 0 Then Secs = 30
    
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg(COLOR_BRIGHTBLUE & "Server Shutdown in " & Secs & " seconds.")
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If
    
    Secs = Secs - 1
    
    If Secs <= 0 Then
        Call GlobalMsg(COLOR_BRIGHTRED & "Server Shutdown.")
        Call DestroyServer
    End If
    
End Sub

